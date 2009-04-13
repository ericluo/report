require 'spreadsheet'

class Aggreator
  attr_accessor :template, :source_pattern, :output, :range

  def initialize(template, source_path, output, range)
    @template = Spreadsheet.open(template)
    @output = output
    source_pattern = File.expand_path(File.join(source_path, "*.xls"))
    @sources = Dir.glob(source_pattern)
    @rows_range, @cols_range = transform(range)
  end

  def transform(range)
    b, e = range.split(':')
    rows_range = (b[1..-1].to_i - 1)..(e[1..-1].to_i - 1)
    cols = ('A'..'Z').to_a
    cols_range = (cols.index(b[0..0]))..(cols.index(e[0..0]))
    [rows_range, cols_range]
  end

  def aggreate
    @sources.each_with_index do |f, i|
      sheet = Spreadsheet.open(f).worksheet(0)
      yield f, i
      @rows_range.each do |r|
        @cols_range.each do |c|
          # debug("r=#{r}, c=#{c}, value=#{sheet.row(r)[c]}")
          @template.worksheet(0).row(r)[c] ||= 0
          @template.worksheet(0).row(r)[c] += sheet.row(r)[c]
        end
      end
    end
    @template.write(@output)
  end

  def size
    @sources.size
  end
end


