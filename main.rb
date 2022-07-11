require 'rubyXL'
require 'rubyXL/convenience_methods/cell'
require 'rubyXL/convenience_methods/color'
require 'rubyXL/convenience_methods/font'
require 'rubyXL/convenience_methods/workbook'
require 'rubyXL/convenience_methods/worksheet'

File.open('./output.xlsx', 'wb') do |file|
  content = RubyXL::Parser.parse('./input.xlsx').tap do |wb|
    sheet = wb.first
    sheet.insert_column(1)
  end
  file.write(content.stream.read)
end