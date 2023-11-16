require 'roo'
require 'spreadsheet'

class Array
  def avg
    sum / size.to_f
  end

  def method_missing(name, *args)
    self.index(name.to_s) ? self.index(name.to_s) : super
  end
end

class ExcelTable
  attr_reader :header, :data

  def initialize(file_path, sheet_name)
    @workbook = Roo::Spreadsheet.open(file_path)
    @sheet = @workbook.sheet(sheet_name)
    @data = (0..@sheet.last_row).map { |row| @sheet.row(row) }

    @data.reject! do |row|
      row.all?(&:nil?) || row.any? { |cell| cell.to_s.downcase.include?('total') || cell.to_s.downcase.include?('subtotal') }
    end
    @header = @data.first
    @data.shift
  end

  def row(index)
    @data[index]
  end

  def each
    yield @header
    @data.each { |row| yield row }
  end

  def [](column_name)
    column_index = @header.index(column_name)
    return nil unless column_index

    @data.map { |row| row[column_index] }
  end

  def []=(column_name, row_index, value)
    column_index = @header.index(column_name)

    return nil unless column_index
    @data[row_index][column_index] = value
  end

  def method_missing(method_name, *args)
    column_name = method_name.to_s
    return super unless @header.include?(column_name)

    column_index = @header.index(column_name)
    @data.map { |row| row[column_index] }
  end

  def array2D
    [@header] + @data
  end

  def +(other_table)
    @header == other_table.header ? @data += other_table.data : puts('Tables have different headers')
  end

  def -(other_table)
    return puts('Tables have different headers') unless @header == other_table.header

    other_table.data.each { |other_row| @data.delete(other_row) }
  end
end