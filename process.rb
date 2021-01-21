#!/usr/bin/env ruby
# frozen_string_literal: true

require 'csv'
require 'spreadsheet'

SPREADSHEET_FILENAME = './inventario.xls'
CSV_OUTPUT = 'inventario.csv'

unless File.exist?(SPREADSHEET_FILENAME)
  puts "Expected file '#{SPREADSHEET_FILENAME}' but it wasn't found.
NOTE: This program only supports .xls files. If you have an .xlsx file, convert it in Excel first."
  exit 1
end

workbook = Spreadsheet.open SPREADSHEET_FILENAME
puts "Processing '#{SPREADSHEET_FILENAME}' ..."
worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets."

def find_headers(worksheet)
  headers = ['Fornecedor']
  headers += worksheet.row(0)
  worksheet.row(1).each_with_index do |value, column|
    next if value.nil?

    headers[column + 1] = value
  end
  headers
end

def define_state(colour)
  colour != :border
end

def before_tax_a_value?(hash)
  !hash['s/ IVA'].is_a? Numeric
end

def description_a_title?(hash)
  hash['Descrição'] =~ /TOTAL|Total|Portes|Desconto|Diversos/
end

def field_matches_subtotal?(field)
  return false if field.is_a? Float

  field =~ /Contab.*/
end

def skip_conditions(hash)
  return true if before_tax_a_value?(hash)

  return true if description_a_title?(hash)

  return true if field_matches_subtotal?(hash['c/ IVA'])

  return true if field_matches_subtotal?(hash['Observações'])

  false
end

inventory = []

def accounting_inventory(inventory)
  CSV.open(CSV_OUTPUT, 'wb', encoding: 'utf-8', col_sep: ';') do |csv|
    csv << [
      'ProductCategory',
      'ProductCode',
      'ProductDescription',
      'ProductNumberCode',
      'ClosingStockQuantity',
      'Unit of Measure',
      'Preço unitário',
      'valor'
    ]
    inventory.each do |hash|
      csv << [
        'M',
        hash['Descrição'],
        hash['Descrição'],
        hash['Descrição'],
        1,
        'UN',
        hash['s/ IVA'],
        hash['s/ IVA']
      ]
    end
  end
end

def normal_inventory(inventory)
  CSV.open(CSV_OUTPUT, 'wb', encoding: 'utf-8', col_sep: ';') do |csv|
    csv << [
      'Fornecedor',
      'Ref Item',
      'Nº Doc.',
      'Descrição',
      'Entrada',
      's/ IVA',
      'c/ IVA',
      'Linha'
    ]
    inventory.each do |hash|
      csv << [
        hash['Fornecedor'],
        hash['Ref Item'] || hash['Referência'],
        hash['Nº Doc.'],
        hash['Descrição'],
        hash['Entrada'] || hash['Data'],
        hash['s/ IVA'],
        hash['c/ IVA'],
        hash['Linha']
      ]
    end
  end
end

worksheets.each do |worksheet|
  puts "Processing '#{worksheet.name}' ..."
  headers = find_headers(worksheet)
  rows = worksheet.rows
  row_count = rows.count - 1
  (2..row_count).each do |row|
    hash = {}

    headers.each_with_index do |header, index|
      contents = rows[row][index - 1]
      hash[header] = contents.respond_to?(:value) ? contents.value : contents
    end
    next if skip_conditions(hash)

    hash['Fornecedor'] = worksheet.name
    hash['Linha'] = row + 1

    # Check row default colour and ignore sold items
    colour = worksheet.row(row).default_format.pattern_fg_color
    next if define_state(colour)

    # Ignore items that have the 5th column set with a background colour
    next if worksheet.row(row).format(5).pattern_bg_color != :pattern_bg

    # Add item to inventory
    inventory << hash
  end
end

puts "Writing output to '#{CSV_OUTPUT}' ..."
accounting_inventory(inventory)
