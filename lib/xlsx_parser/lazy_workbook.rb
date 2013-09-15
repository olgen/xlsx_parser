module XlsxParser
  class LazyWorkbook < Workbook

    # is a bit arbitrary atm
    CHUNK_SIZE = 256

    def initialize(filename)
      @tmpdir = Dir.mktmpdir
      @tmp_file_name = copy_file_to_tmp_dir(filename)
      extract_zip_content(@tmp_file_name)
      read_workbook
      parse_shared_strings
    end

    # incrementally read the sheet and yield the rows, to save memory
    # Use the PushParser, which we feed manually, so we can modify the amount of parsed nodes
    def each_row(sheet, max_rows = nil, &block)
      sax_document = LazySheetSaxParser.new(@shared_strings, &block)
      parser = Nokogiri::XML::SAX::PushParser.new(sax_document)

      io = sheet_io(sheet)
      while chunk = io.read(CHUNK_SIZE)
        parser << chunk # feed parser, expect callbacks
        break if max_rows && sax_document.rows_loaded >= max_rows
      end
      io.close
    end

    private

    def sheet_io(sheet)
      sheet_file = @sheet_files[sheets.index(sheet)]
      File.open(sheet_file)
    end

  end
end

