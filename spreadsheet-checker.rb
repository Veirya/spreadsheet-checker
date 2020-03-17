=begin
  Spreadsheet Checker, by Giovanni "Veirya" Oliver (2020)
  :::::::::::::::::::::::::::::::::::::::::::::::::
  Simple script that looks at a Google Sheets sheet of FFXIV FC member data,
    then compares it to information from the Lodestone. Information considered
    to be of interest is then printed out for the user to see. Current iteration
    only works for the FC "FullMetal Alliance" and a hard-coded sheet.
  Information considered of interest by the script:
    - New/Old Members
    - Name Changes
    - Rank Changes
    - Possibile Promotion Candidates (FMA New Recruits only)
    - Lack of Date for a Member in Spreadsheet.
=end

# Dependencies
#####################
require 'net/http'
require 'json'
require 'ostruct'
require 'date'
require "google/apis/sheets_v4"
require "googleauth"
require "googleauth/stores/file_token_store"
require "fileutils"

# Constants
#####################
OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "FC Spreadsheet Checker".freeze
CREDENTIALS_PATH = "credentials.json".freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY
TODAY = DateTime.parse(Time.now.strftime("%y-%m-%d"))

# Methods
#####################

##
# Ensure valid credentials, either by restoring from the saved credentials
# files or intitiating an OAuth2 authorization. If authorization is required,
# the user's default browser will be launched to approve the request.
#
# @return [Google::Auth::UserRefreshCredentials] OAuth2 credentials
def authorize
  client_id = Google::Auth::ClientId.from_file CREDENTIALS_PATH
  token_store = Google::Auth::Stores::FileTokenStore.new file: TOKEN_PATH
  authorizer = Google::Auth::UserAuthorizer.new client_id, SCOPE, token_store
  user_id = "default"
  credentials = authorizer.get_credentials user_id
  if credentials.nil?
    url = authorizer.get_authorization_url base_url: OOB_URI
    puts "Open the following URL in the browser and enter the " \
         "resulting code after authorization:\n" + url
    code = gets
    credentials = authorizer.get_and_store_credentials_from_code(
      user_id: user_id, code: code, base_url: OOB_URI
    )
  end
  credentials
end

##
# Reads and formats member data from a Google Sheets spreadsheet.
#
# [Google::Apis::SheetsV4::SheetsService] service: Google Sheets API service
# @return [Hash[Integer : OpenStruct]] Formatted member data from sheet 
def readSheet(service)
  print "Retrieving FC members from spreadsheet... "
  # Retrieve Current FC members data from spreadsheet
  spreadsheet_id = "1i09Ey3KFzvJENkToV1o5bSh89AML85fW6cz0o0IMgI0"
  range = "Members (Condensed)!A2:I"
  response = service.get_spreadsheet_values spreadsheet_id, range

  sheetMembers = Hash.new
  response.values.each do |row|
    data = OpenStruct.new
    data.Name = row[0]
    data.Rank = row[2]

    # Dates should be in format of MM/DD/YYYY. Empty entries will be flagged.
    data.Date = row[5]
    if data.Date == ""
      data.Date = row[4]
    end
    if data.Date != ""
      data.Date = data.Date.split("/")
      data.Date = DateTime.parse("#{data.Date[2]}-#{data.Date[0]}-#{data.Date[1]}")
    end

    data.ID = row[8].to_i
    sheetMembers[data.ID] = data
  end
  puts "Finished."

  return sheetMembers
end

##
# Retrieves and formats member data from the FFXIV Lodestone.
#
# @return [Hash[Integer : OpenStruct]] Formatted member data from Lodestone
def checkLode()
  print "Retrieving FC members from xivapi... "
  # Retrieve current FC members from Lodestone and convert to hash
  # mapped by ID
  url = "https://xivapi.com/freecompany/9232097761132854687?data=FCM"
  uri = URI(url)
  response = Net::HTTP.get(uri)
  parse = JSON.parse(response, object_class: OpenStruct).FreeCompanyMembers
  lodeMembers = Hash.new

  parse.each do |member|
  	lodeMembers[member.ID] = member
  end
  puts "Finished."
  return lodeMembers
end

##
# Compare member data from sheet and Lodestone, flagging information considered to be
# of interest to the user. May remove entries from sheetMembers. Does not modify
# lodeMembers. Fills requested parameters with their respective flagged information.
#
# [Hash[Integer : OpenStruct]] sheetMembers : Member data from a Google Sheets spreadsheet
# [Hash[Integer : OpenStruct]] lodeMembers : Member data from the FFXIV Lodestone
# [Array] <<An empty Array for each type of information of interest>>
# @return nil
def processData(sheetMembers, lodeMembers, newMems, nameChanges, rankChanges, noDates, needPromos)
  lodeMembers.each do |key, member|
  	if sheetMembers[key]
  		if sheetMembers[key].Name != member.Name
  			nameChanges.push("'#{sheetMembers[key].Name}' changed their name to '#{member.Name}'.")
  		end

      if sheetMembers[key].Rank != member.Rank
        rankChanges.push("#{sheetMembers[key].Name} has changed ranks from ".concat(
          "'#{sheetMembers[key].Rank}' to '#{member.Rank}'."))
      end

      if sheetMembers[key].Date.is_a? String
        noDates.push("#{member.Name}")
      elsif member.Rank == "New Recruit"
        diff = (TODAY - sheetMembers[key].Date).to_i
        if diff >= 30
          needPromos.push("#{member.Rank} #{member.Name} may need a promotion. (#{diff} days since joined)")
        end
      end

  		sheetMembers.delete(key)
  	else
      # New members.
  		newMems.push("#{member.Name}\tID: #{key}")
  	end
  end
  return nil
end

##
# Prints the data from the requested parameters in a user-readable format.
#
# [Hash[Integer : OpenStruct]] sheetMembers : Processed member data
# [Array] <<Arrays of strings containing information of interest>>
# @return nil
def printData(sheetMembers, newMems, nameChanges, rankChanges, noDates, needPromos)
  puts "New Members"
  puts ":::::::::::::::::::::::"
  newMems.each do |string|
  	puts string
  end
  puts ""

  puts "No longer in FC"
  puts ":::::::::::::::::::::::"
  sheetMembers.each do |key, member|
  	puts "#{member.Name}\tID: #{key}"
  end
  puts ""

  puts "Name Changes"
  puts ":::::::::::::::::::::::"
  nameChanges.each do |string|
  	puts string
  end
  puts ""

  puts "Rank Changes"
  puts ":::::::::::::::::::::::"
  rankChanges.each do |string|
    puts string
  end
  puts ""

  puts "Possibile Promotion Candidates"
  puts ":::::::::::::::::::::::"
  needPromos.each do |string|
    puts string
  end
  puts ""

  puts "In Need of Dates"
  puts ":::::::::::::::::::::::"
  noDates.each do |string|
    puts string
  end
  puts ""
  return nil
end

# Main Code Body
#####################
puts "Today's date is #{TODAY.mon}/#{TODAY.day}/#{TODAY.year}."
# Initialize the Google Sheets API
print "Authenticating and Initializing Google API... "
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize
puts "Finished."

sheetMembers = readSheet(service)
lodeMembers = checkLode()
puts ""

newMems = Array.new
nameChanges = Array.new
rankChanges = Array.new
noDates = Array.new
needPromos = Array.new

processData(sheetMembers, lodeMembers, newMems, nameChanges, rankChanges, noDates, needPromos)
printData(sheetMembers, newMems, nameChanges, rankChanges, noDates, needPromos)

print "Press enter to quit."
gets.chomp