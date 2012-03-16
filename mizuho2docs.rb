#!/usr/bin/ruby -Ku
require "rubygems"
require 'yaml'
require 'time'
require_relative "mizuhodirect/mizuhodirect"
require_relative 'googlespreadsheet/spreadsheet'

mizuho_account = YAML.load_file(File.dirname(__FILE__) + '/mizuho_account.yaml')
google_account = YAML.load_file(File.dirname(__FILE__) + '/account.yaml')


spreadsheet_key = google_account['spreadsheet_key']

# spreadsheet
session = GoogleSpreadsheet.login(google_account['email'], google_account['passwd'])
ws = session.spreadsheet_by_key(spreadsheet_key).worksheets[0]

# find last (FIXME!
last = ws[ws.row_count]
unless last[0]
  last = ws[ws.row_count-1]
end


# login
m = MizuhoDirect.new
unless m.login(mizuho_account)
  puts "LOGIN ERROR"
end

begin
  account_status = m.get_top

  puts 'total: ' + account_status["zandaka"].to_s

  st = nil
  account_status["recentlog"].each{|row|
    if Time.parse(last[0])==Time.parse(row[0]) && last[1].to_i == row[1] && last[2].to_i == row[2]
      st = row
      #puts "match!"
      #p row
    end
  }

  account_status["recentlog"].each do |row|
    if st
      if row == st
        st = nil
      end
      next
    end
    p row
    ws << row
  end
  ws[ws.row_count,5] = account_status["zandaka"]

ensure
  # logout
  m.logout
end

puts "ok"

