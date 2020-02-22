require 'net/http'
require 'json'
require 'ostruct'
require "google/apis/sheets_v4"
require "googleauth"
require "googleauth/stores/file_token_store"
require "fileutils"

OOB_URI = "urn:ietf:wg:oauth:2.0:oob".freeze
APPLICATION_NAME = "Google Sheets API Ruby Quickstart".freeze
CREDENTIALS_PATH = "credentials.json".freeze
# The file token.yaml stores the user's access and refresh tokens, and is
# created automatically when the authorization flow completes for the first
# time.
TOKEN_PATH = "token.yaml".freeze
SCOPE = Google::Apis::SheetsV4::AUTH_SPREADSHEETS_READONLY

print "Authenticating and Initializing Google API... "
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

# Initialize the API
service = Google::Apis::SheetsV4::SheetsService.new
service.client_options.application_name = APPLICATION_NAME
service.authorization = authorize
puts "Finished."

print "Retrieving FC members from spreadsheet... "
# Retrieve Current FC members from spreadsheet and convert to hash
# mapped by ID
spreadsheet_id = "1i09Ey3KFzvJENkToV1o5bSh89AML85fW6cz0o0IMgI0"
range = "Members (Condensed)!A2:I"
response = service.get_spreadsheet_values spreadsheet_id, range

sheetMembers = Hash.new
response.values.each do |row|
  data = OpenStruct.new
  data.Name = row[0]
  data.Rank = row[2]
  data.ID = row[8].to_i
  sheetMembers[data.ID] = data
end
puts "Finished."

print "Retrieving FC members from xivapi... "
# Retrieve current FC members from Lodestone and convert to hash
# mapped by ID
url = "https://xivapi.com/freecompany/9232097761132854687?data=FCM"
uri = URI(url)
response = Net::HTTP.get(uri)
parse = JSON.parse(response, object_class: OpenStruct).FreeCompanyMembers
lodeMembers = Hash.new

parse.each do |member|
	member.ID = 
	lodeMembers[member.ID] = member
end
puts "Finished."
puts ""

=begin
		
=end
newMems = Hash.new
nameChanges = Array.new
lodeMembers.each do |key, member|
	if sheetMembers[key]
		if sheetMembers[key].Name != member.Name
			nameChanges.push("'#{sheetMembers[key].Name}' changed their name to '#{member.Name}'.")
		end
		sheetMembers.delete(key)
	else
		newMems[key] = member
	end
end

puts "New Members:"
newMems.each do |key, member|
	puts "Name: #{member.Name}, ID: #{key}"
end
puts ""

puts "No longer in FC:"
sheetMembers.each do |key, member|
	puts "Name: #{member.Name}, ID: #{key}"
end
puts ""

puts "Name Changes:"
nameChanges.each do |string|
	puts string
end
puts ""

print "Press enter to quit."
gets.chomp