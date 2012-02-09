require 'open3'
# TODO: clipboard and encoding .. what locale is it?
module Clipboard; end

module Clipboard::Windows
  extend self

  CF_TEXT = 1
  CF_OEMTEXT = 7
  CF_UNICODETEXT = 13
  GMEM_MOVEABLE = 2
  
  CF_HDROP = 15 # file references
    
  $text_formats = [CF_TEXT, CF_OEMTEXT, CF_UNICODETEXT, 'text/plain', 'text/html' ] 
  TYPEMAP = Hash[ 'text/html', 'HTML format', 'text/plain', CF_UNICODETEXT, 'file reference', CF_HDROP ]

  # get ffi function handlers
  begin
    require 'ffi'
  rescue LoadError
    raise LoadError, 'Could not load the required ffi gem, install it with: gem install ffi'
  end

  module User32
    extend FFI::Library
    ffi_lib "user32"
    ffi_convention :stdcall

    attach_function :open,  :OpenClipboard,    [ :long ], :long
    attach_function :close, :CloseClipboard,   [       ], :long
    attach_function :empty, :EmptyClipboard,   [       ], :long
    attach_function :get,   :GetClipboardData, [ :long ], :long
    attach_function :set,   :SetClipboardData, [ :long, :long ], :long
    attach_function :register,   :RegisterClipboardFormatA, [ :string ], :long
  end
  CF_HTML = User32.register('HTML Format')
 # $text_formats << CF_HTML
  module Kernel32
    extend FFI::Library
    ffi_lib 'kernel32'
    ffi_convention :stdcall

    attach_function :lock,   :GlobalLock,   [ :long ], :pointer
    attach_function :unlock, :GlobalUnlock, [ :long ], :long
    attach_function :alloc,  :GlobalAlloc,  [ :long, :long ], :long
    attach_function :size,  :GlobalSize,  [ :long ], :long
  end

  # see http://www.codeproject.com/KB/clipboard/archerclipboard1.aspx
  def paste( format = CF_UNICODETEXT)
    ret = ""
      if 0 != User32.open( 0 )
        win_format = get_win_format_code format
        hclip = User32.get( win_format )
        if hclip && 0 != hclip
          pointer_to_data = Kernel32.lock( hclip )
          size = Kernel32.size hclip # Not sure if GlobalSize will work for all sorts of clipboard contents.. :-o
          data = pointer_to_data.get_bytes( 0, size-1 )
          ret = data
          if $text_formats.include?(format)
            if data.encoding # Ruby 1.9 and greater
              begin
                Encoding.default_external 'UTF-8'
                ret = data.chop.force_encoding("UTF-16LE").encode(Encoding.default_external)
              rescue
              end
            else # 1.8: fallback to simple CP850 encoding
              require 'iconv'
              utf8 = Iconv.iconv( "UTF-8", "UTF-16LE", data.chop)[0]
              ret = Iconv.iconv( "CP850", "UTF-8", utf8)[0]
            end
        end
        if data && 0 != data
          Kernel32.unlock( hclip )
        end
      end
      User32.close( )
    end
    if format.eql? 'text/html'
      startF = ret.match(/StartFragment:(\d+)/)[1].to_i
      endF = (ret.match(/EndFragment:(\d+)/)[1].to_i) -1
      ret = ret[startF..endF]
    end
    if format.eql? 'file reference'
      ret = data.chop.force_encoding("UTF-16LE").encode('UTF-8')
      ret = ret.match('\x01\x00(.*)\000$') [1].split("\x00")
      if ret.length==1
        ret=ret[0]
      end
    end
    ret || ""
  end

  def clear
    if 0 != User32.open( 0 )
      User32.empty( )
      User32.close( )
    end
    paste
  end

  def copy(data_to_copy, format = 'text/plain', alternate=false)
    if ( "".encoding ) && 0 != User32.open( 0 )
      win_format = get_win_format_code format
      if $text_formats.include? format then
        data = data_to_copy.encode("UTF-16LE") # TODO catch bad encodings
        data << 0 << 0
      end
      if format.eql? 'text/html' # Make sure there is an alternate part with text/plain
        # This is of course a very crude way to strip tags. Things like hello<br>world won't come out correctly..
        textplain = data_to_copy.gsub(/<\/?[^>]*>/, "")
        data_to_copy = construct_html_format data_to_copy
      end
      data=data_to_copy
      if format.eql? 'file reference' then # create (eh, fake) Windows file ref structure
        if data_to_copy.kind_of? Array then 
          data_to_copy = data_to_copy.join "\0"
        end
        data=data_to_copy.force_encoding( 'UTF-8' ).encode "UTF-16LE"
        data = ("\024"+( "\000".*7 )+"\001"+"\000").force_encoding( 'UTF-8' ).encode("UTF-16LE")+data
        data <<0 <<0
      end
      # TODO: what about binary data??
      if ! alternate then
        User32.empty( )
      end
      handler = Kernel32.alloc( GMEM_MOVEABLE, data.bytesize )
      pointer_to_data = Kernel32.lock( handler )
      pointer_to_data.put_bytes( 0, data, 0, data.bytesize )
      Kernel32.unlock( handler )
      User32.set( win_format, handler )
      User32.close( )
      # If we put text/html on the clipboard, we should include a text/plain alternative.
      if format.eql? 'text/html' then
        copy textplain, 'text/plain', true
      end
    else # don't touch anything
      Open3.popen3( 'clip' ){ |input,_,_| input << data_to_copy } # depends on clip (available by default since Vista)
    end
    paste
  end
  def get_win_format_code( format )
    if is_numeric?(format)
      return format;
    end
    if is_numeric?( TYPEMAP[format] ) 
      return TYPEMAP[format];
    end
    win_form = User32.register( TYPEMAP[format] )
    if win_form != 0 
      return win_form
    end
  end
  def is_numeric?(input)
    begin Float(input) ; true end rescue false
  end
  def construct_html_format ( html )
    metablock = %{Version:0.9
StartHTML:
EndHTML:
StartFragment:
EndFragment:
<!DOCTYPE>
<html>
<head>
<title></title>
</head>
<body>
<!--StartFragment -->
<!--EndFragment -->
</body>
</html>}
    metablock = metablock.gsub /\n/, "\r\n"
    metablock[/<!--EndFragment -->/]=html+"\r\n<!--EndFragment -->"
    starthtml = metablock.index '<!DOCTYPE'
    startfragment = metablock.index( '<!--StartFragment -->' )+21
    endfragment = metablock.index '<!--EndFragment -->' 
    endhtml = metablock.length
    # meta data has to include size of meta data itself..
    additional_sizes=starthtml.to_s.size + startfragment.to_s.size+endfragment.to_s.size + endhtml.to_s.size
    # for each of those that will be one digit longer when size is added, we must add some more..
    if (starthtml+additional_sizes).to_s.size > starthtml.to_s.size 
      additional_sizes +=  (starthtml+additional_sizes).to_s.size - starthtml.to_s.size 
    end
    if (startfragment+additional_sizes).to_s.size > startfragment.to_s.size 
      additional_sizes +=  (startfragment+additional_sizes).to_s.size - startfragment.to_s.size 
    end
    if (endfragment+additional_sizes).to_s.size > endfragment.to_s.size 
      additional_sizes +=  (endfragment+additional_sizes).to_s.size - endfragment.to_s.size 
    end
    if (endhtml+additional_sizes).to_s.size > endhtml.to_s.size 
      additional_sizes +=  (endhtml+additional_sizes).to_s.size - endhtml.to_s.size 
    end
    metablock[ /StartHTML:/ ] = 'StartHTML:'+(starthtml+additional_sizes).to_s
    metablock[ /StartFragment:/ ] = 'StartFragment:'+(startfragment +additional_sizes).to_s 
    metablock[ /EndFragment:/ ] = 'EndFragment:'+(endfragment +additional_sizes).to_s 
    
    metablock[ /EndHTML:/ ] = 'EndHTML:'+ metablock.length.to_s
    
    return metablock
  end
end
