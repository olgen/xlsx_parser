module XlsxParser
  class LazySheetSaxParser < SheetSaxParser
    attr_reader :rows_loaded

    def initialize(shared_strings, &block)
      raise "Need some shared strings to populate the sheet!" unless shared_strings
      raise "Need a block for row callbacks" unless block
      @shared_strings = shared_strings
      @block = block
      @buffer = ''
      @rows_loaded = 0
    end

    def start_row
      @current_row = []
    end

    # yield after every row
    def end_row
      @block.call(@current_row)
      @rows_loaded += 1
    end

  end
end
