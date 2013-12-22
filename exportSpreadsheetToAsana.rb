#!/usr/bin/env ruby
 
require "rubygems"
require 'cgi'
require 'google_drive'
require 'asana'
require 'yaml'


cnf = YAML::load(File.open('config.yml'))

# Google drive keys
GOOGLE_DRIVE_USERNAME = cnf['google_drive']['username'] 
GOOGLE_DRIVE_PASSWORD = cnf['google_drive']['password'] 

# Asana keys
ASANA_API_KEY = cnf['asana']['api_key'] 
ASANA_ASSIGNEE = 'me'
 

def get_option_from_list(list, title, attribute)
  i=0

  while i == 0 do
    puts title
    
    list.each do |item|
      i += 1
      puts "  #{i}) #{item.send(attribute)}"
    end

    i = gets.chomp.to_i
    i = 0 if i <= 0 && i > list.size    
  end
  return i - 1
end
 
Asana.configure do |client|
  client.api_key = ASANA_API_KEY
end

workspaces = Asana::Workspace.all

session = GoogleDrive.login(GOOGLE_DRIVE_USERNAME, GOOGLE_DRIVE_PASSWORD)

puts "What's the spreadsheet url?"
u = URI.parse(gets.chomp)
params = CGI.parse(u.query)
raise 'Could not find key param in url.' unless params.has_key?('key')
sheetKey = params['key'].first

ws = session.spreadsheet_by_key(sheetKey).worksheets[0]
puts 'Got the spreadsheet successfully'

puts "Num rows: #{ws.num_rows}"

# Which workspace to put it in
workspace = workspaces[get_option_from_list(workspaces, 
  "Select destination workplace", 
  "name")]
puts "Using workspace #{workspace.name}"

# Which project to associate
project = workspace.projects[get_option_from_list(workspace.projects, 
  "Select destination project", 
  "name")]
puts " -- Using project #{project.name} --"


ws.num_rows.downto(1) do |row|
  if !ws[row, 2].empty?
    puts "  #{ws[row,2]}, Due on #{ws[row,4]}"
    t = Asana::Task.new
    t.name = ws[row,2]
    t.notes = ws[row,6] if !ws[row,6].empty?
    t.assignee = nil
    begin
      t.due_on = DateTime.strptime(ws[row,4],"%m/%d/%Y") if !ws[row,4].empty?
    rescue
    end
    task = workspace.create_task(t.attributes)
    task.add_project(project.id)
  end
  if !ws[row, 1].empty? && ws[row, 1].downcase!=ws[row, 2].downcase
    puts "#{ws[row,1]}:"
    t = Asana::Task.new
    t.name = "#{ws[row,1]}:"
    t.assignee = nil
    task = workspace.create_task(t.attributes)
    task.add_project(project.id)
  end
end

