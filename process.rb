#!/usr/bin/env ruby
# frozen_string_literal: true

require 'csv'
require 'spreadsheet'

# Note: spreadsheet only supports .xls files (not .xlsx)
workbook = Spreadsheet.open './inv.xls'
worksheets = workbook.worksheets
puts "Found #{worksheets.count} worksheets"

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

def skip_conditions(hash)
  (!hash['s/ IVA'].is_a? Numeric) || hash['Descrição'] =~ /TOTAL|Total|Portes|Desconto|Diversos/ || hash['c/ IVA'] =~ /Contab.*/ || hash['Observações'] =~ /Contab.*/
end

inventory = []

def print_inventory(inventory)
  CSV.open('tmp.csv', 'wb', encoding: 'utf-8:ISO8859-1') do |csv|
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

print_inventory(inventory)
