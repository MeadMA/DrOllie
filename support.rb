require 'win32/registry'
require 'json'
require 'open-uri'
require 'Win32API'
require 'fileutils'
require 'net/http'

# Query the 'uninstall' key from the Registry to get a list of installed
# software.  Return an array of arrays.  Each member array is the name of the
# software and its version.
# If "true" is passed to this method, query from Wow6432Node (32-bit software
# installed on 64-bit system)
def query_uninstall(wow64=false)
	if wow64
		key = 'Software\Wow6432Node\Microsoft\Windows\CurrentVersion\Uninstall'
	else
		key = 'Software\Microsoft\Windows\CurrentVersion\Uninstall'
	end
	ary = []
	# Iterate through each key under Uninstall
	Win32::Registry::HKEY_LOCAL_MACHINE.open(key) do |reg|
		reg.each_key do |key1,key2|
			k = reg.open(key1)
			name = ''
			version = ''
			# Get the DisplayName value; use rescue in case the value doesn't
			# exist
			begin
				name = k["DisplayName"]
			rescue
			end
			# Get the DisplayVersion value; use rescue in case the value doesn't
			# exist
			begin
				version = k["DisplayVersion"]
			rescue
			end
			ary << [name, version]
		end
	end
	# Delete entries where the DisplayName value was blank or didn't exist
	ary.delete_if do |n, v|
		n == ''
	end
	# Return the collected software
	return ary
end

# Read the .CONF JSON files from the 'def' directory and return an array.
# Each entry in the array is a hash representing a definition.
def read_defs
	defs = Array.new
	# Get list of .CONF files from 'def' directory
	def_files = Dir["#{DATA_DIR}/sync/def/*.conf"]
	# Read each file and add it to the array
	def_files.each do |file|
		f = File.read(file)
		d = JSON.parse(f)
		# Convert the "name" string to a regular expression
		d["name"] = Regexp.new d["name"]
		# Convert the "success_codes" string to an array
		d["success_codes"] = d["success_codes"].split(',')
		defs << d
	end
	# Return the array of definitions
	return defs
end

# "installed_program" method
# Given a definition and an array of installed software (collected using
# "query_uninstall" method), check if software is outdated
# Returns 'true' if outdated and 'false' if not
def installed_program_outdated?(def_hsh, software)
	# Iterate through each item of installed software
	software.each do |name, version|
		# If the software's name matches the regular expression defined in the
		# definition, then compare versions
		if name =~ def_hsh["name"]
			# Use "version_less_than?" method to see if installed software's
			# version is less than the version from the definition
			if version_less_than?(version, def_hsh["min_version"])
				log_event(7, 'INFORMATION', "#{def_hsh["description"]} is out-of-date.")
				return true
			end
		end
	end
	return false
end

def file_version_outdated?(def_hsh)
	current_ver = get_file_version(def_hsh["check_file"])
	if version_less_than?(current_ver, def_hsh["min_version"])
		log_event(7, 'INFORMATION', "#{def_hsh["description"]} is out-of-date.")
		return true
	end
	return false
end

# "registry_version" method
# Given a definition, grab a version from the Registry value specified
# Return true if outdated, false if not
def registry_version_outdated?(def_hsh)
	case def_hsh["registry_hive"]
	when "HKLM", "HKEY_LOCAL_MACHINE"
		keycmd = reg_hklm
	when "HKCU", "HKEY_CURRENT_USERS"
		keycmd = reg_hkcu
	when "HKU", "HKEY_USERS"
		keycmd = reg_hku
	when "HKCC", "HKEY_CURRENT_CONFIG"
		keycmd = reg_hkcc
	when "HKCR", "HKEY_CLASSES_ROOT"
		keycmd = reg_hkcr
	else
		log_event(9, 'ERROR', "The registry hive specified ('#{def_hsh["registry_hive"]}' in the definition for '#{def_hsh["description"]}' is invalid.")
		return false
	end
	
	reg_version = send(keycmd, def_hsh["registry_key"], def_hsh["registry_value"])
	unless reg_version
		return false
	end
	if version_less_than?(reg_version, def_hsh["min_version"])
		log_event(7, 'INFORMATION', "#{def_hsh["description"]} is out-of-date.")
		return true
	end
	return false
end

# Given a regular expression and an array of installed software, see if any of
# the installed software matches the expression
# Return 'true' if installed or 'false' if not
def is_installed?(name_regexp, software)
	software.each do |name, version|
		if name =~ name_regexp
			return true
		end
	end
	return false
end

# Downloads from a specified URL to a specified destination
def download_update(url, dest)
	if File.exist?(dest)
		log_event(2, 'INFORMATION', "Download of '#{url}' to '#{dest}' cancelled. '#{dest}' already exists.")
	else
		log_event(1, 'INFORMATION', "Downloading '#{url}' to '#{dest}'.")
		open(dest, 'wb') do |file|
			file << open(url).read
		end
	end
end

# Given two versions, see if the first version is smaller/older than the second
# version
# Return true if smaller/older, false if not
def version_less_than?(ver1, ver2)
	# Some sick manufacturers use commas instead of periods in their versions
	# Convert these commas to periods for the comparison
	ver1 = ver1.gsub(',', '.')
	ver2 = ver2.gsub(',', '.')
	# Split the versions at each period
	ver1 = ver1.split('.')
	ver2 = ver2.split('.')
	
	# For each level, see if the first version is less than the second version
	# Example, 5.0.6.4 and 5.0.6.8 will get caught on the fourth level
	for i in 0..9
		if ver1[i].to_i < ver2[i].to_i
			return true
		end
	end
	return false
end

# Use 'eventcreate' command to log an event in the Application log, using the
# source name "DrOllie".  Use the provided ID, event type (SUCCESS, ERROR,
# WARNING, or INFORMATION), and description.
# Also print the event to the command line
def log_event(id, type, desc)
	print "#{type}: #{desc}\n"
	desc = desc.gsub("\n", "  ")
	`eventcreate /l APPLICATION /t #{type.upcase} /so DrOllie /id #{id.to_s} /d "#{desc}"`
end

# Reads a given key/value from HKLM
def reg_hklm(key, value)
	begin
		Win32::Registry::HKEY_LOCAL_MACHINE.open(key) do |reg|
			return reg[value]
		end
	rescue
		return false
	end
end

# Reads a given key/value from HKCU
def reg_hkcu(key, value)
	begin
		Win32::Registry::HKEY_CURRENT_USER.open(key) do |reg|
			return reg[value]
		end
	rescue
		return false
	end
end

# Reads a given key/value from HKU
def reg_hku(key, value)
	begin
		Win32::Registry::HKEY_USERS.open(key) do |reg|
			return reg[value]
		end
	rescue
		return false
	end
end

# Reads a given key/value from HKCC
def reg_hkcc(key, value)
	begin
		Win32::Registry::HKEY_CURRENT_CONFIG.open(key) do |reg|
			return reg[value]
		end
	rescue
		return false
	end
end

# Reads a given key/value from HKCR
def reg_hkcr(key, value)
	begin
		Win32::Registry::HKEY_CLASSES_ROOT.open(key) do |reg|
			return reg[value]
		end
	rescue
		return false
	end
end

# Attempts to query the user Registry hive for the local SYSTEM account. If it
# fails, the user is either not an administrator or is not UAC-elevated.
# Returns true if an admin or false if not.
def is_admin?
	is_admin = false
	begin
		Win32::Registry::HKEY_USERS.open('S-1-5-19') { |reg| }
		is_admin = true
	rescue
	end
	return is_admin
end

# Perform the update action specified by the definition
# Print output to a log in the "output" directory
def update_action(def_hsh)
	# Set command based on action specified in the definition
	case def_hsh["action"].downcase
	when "execute_downloaded_file"
		if def_hsh["action_parameters"]
			cmd = "\"#{DATA_DIR}/update_files/#{def_hsh["file_name"]}\" #{def_hsh["action_parameters"]}"
		else
			cmd = "\"#{DATA_DIR}/update_files/#{def_hsh["file_name"]}\""
		end
	end
	
	# Execute the command determined above
	log_event(3, 'INFORMATION', "Executing update action for #{def_hsh["description"]} (#{cmd})")
	output = `#{cmd}`
	exit_code=$?.exitstatus.to_s
	
	# Log output from command; convert line feeds to Windows style (yuck)
	log_file = "#{DATA_DIR}/output/#{def_hsh["action"]}_#{Time.now.strftime("%Y%m%d_%H%M%S")}.log"
	output = output.gsub("\r\n", "\n").gsub("\n", "\r\n")
	File.open(log_file, 'w') do |file|
		file.puts output
	end
	
	# Check exit code against successful codes from the definition
	# If succeeded, delete the update file
	if def_hsh["success_codes"].include?(exit_code)
		log_event(4, 'SUCCESS', "Update action for #{def_hsh["description"]} succeeded with exit code #{exit_code}.")
		File.delete("#{DATA_DIR}/update_files/#{def_hsh["file_name"]}")
		return true
	else
		log_event(5, 'WARNING', "Update action for #{def_hsh["description"]} failed with exit code #{exit_code}.")
		return false
	end
end

# Initiates the update sequence for a given definition
def init_update(def_hsh)
	# Get update URL from definition
	url = def_hsh["download_url"]
	# Download to 'update_files' directory and name the file based on
	# the definition
	dest = "#{DATA_DIR}/update_files/" + def_hsh["file_name"]
	# Perform the download
	download_update(url, dest)
	# Execute update action and disable program's auto-update if successful
	if update_action(def_hsh)
		disable_auto_update(def_hsh)
	end
end

# Disables a program's own auto-update function using the script/program
# specified in the definition
def disable_auto_update(def_hsh)
	if def_hsh["disable_auto_update"]
		cmd = "\"#{DATA_DIR}/sync/disable_auto_update/#{def_hsh['disable_auto_update']}\""
		`#{cmd}`
		log_event(10, 'INFORMATION', "Disabled manufacturer's auto-update mechanism for '#{def_hsh['description']}' using '#{cmd}'")
	end
end

def get_file_version(filename)
	s=""
	vsize=Win32API.new('version.dll', 'GetFileVersionInfoSize', ['P', 'P'], 'L').call(filename, s)
	if (vsize > 0)
		result = ' '*vsize
		Win32API.new('version.dll', 'GetFileVersionInfo', ['P', 'L', 'L', 'P'], 'L').call(filename, 0, vsize, result)
		rstring = result.unpack('v*').map{|s| s.chr if s<256}*''
		r = /FileVersion..(.*?)\000/.match(rstring)
		if r
			return r[1]
		else
			return false
		end
	else
		return false
	end
end

# Returns directory for Dr. Ollie data (configuration, definitions, etc.)
def get_data_dir
	key = "SOFTWARE\\Microsoft\\Windows\\CurrentVersion\\Explorer\\Shell Folders"
	appdata = ""
	begin
		Win32::Registry::HKEY_LOCAL_MACHINE.open(key) do |reg|
			appdata = reg['Common AppData'].gsub("\\", "/")
		end
	rescue
		appdata = ENV['SystemDrive'].gsub("\\", "/") + "/ProgramData"
	end
	return "#{appdata}/DrOllie"
end

# Checks if necessary directories exist and creates any that are missing
def ensure_dirs
	[ 'output', 'update_files' ].each do |dir|
		dir = DATA_DIR + "/" + dir
		unless File.exist?(dir)
			unless FileUtils::mkdir_p dir
				log_event(10, 'ERROR', "Failed to create a required directory: '#{dir}'")
				exit(10)
			end
		end
	end
end

# Read the core configuration from DrOllie.conf
def read_core_conf
	filename = DATA_DIR + "/DrOllie.conf"
	unless File.exist?(filename)
		log_event(12, 'ERROR', "DrOllie.conf does not exist or could not be read")
		exit(12)
	end
	file = File.read(filename)
	return JSON.parse(file)
end

# Read the status from status.conf
def read_status
	filename = DATA_DIR + "/status.conf"
	if File.exist?(filename)
		file = File.read(filename)
		return JSON.parse(file)
	else
		return Hash.new
	end
end

# Update status.conf
def save_status(hsh)
	File.write("#{DATA_DIR}/status.conf", hsh.to_json)
end

# Updates definitions using SVN
def update_definitions_svn
	sync_dir = DATA_DIR + "/sync"
	if File.exist?("#{sync_dir}/.svn")
		`svn update "#{sync_dir}"`
	else
		`svn checkout #{CORE_CONF['definitions_url']} "#{sync_dir}"`
	end
end

# Set global variables
DATA_DIR=get_data_dir
CORE_CONF=read_core_conf
STATUS=read_status