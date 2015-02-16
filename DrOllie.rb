load 'support.rb'

# Warn and exit if user is not an administrator or UAC-elevated.
unless is_admin?
	msg = "ERROR: Dr. Ollie must be run with administrative credentials.\nEnsure the user is an administrator and UAC-elevated.\n"
	log_event(13, 'ERROR', msg)
	exit(13)
end

# Perform check for necessary directories
ensure_dirs

# Download definitions and install if newer
update_definitions_svn

# Determine architecture
if RbConfig::CONFIG['host_cpu'] == 'x86_64'
	arch = 64
else
	arch = 32
end

# If 64-bit, get 64-bit software and 32-bit software (from WOW6432Node).  If
# 32-bit, just get 32-bit software.
if arch == 64
	software32 = query_uninstall(true)
	software64 = query_uninstall
else
	software32 = query_uninstall
	software64 = []
end

# Read in definitions from .CONF files in 'def' directory.
defs = read_defs

# Run through each definition and run the check configured.
defs.each do |d|
	# Skip 64-bit checks when running 32-bit
	if d["arch"] == "64" && arch == 32
		next
	end
	
	log_event(6, 'INFORMATION', "Running check for #{d["description"]}")
	# Read architecture defined for this check and grab applicable software for
	# check.
	if d["arch"] == "64"
		software = software64
	else
		software = software32
	end
	# Perform check based on the "method" defined
	case d["method"]
	when "installed_program"
		if installed_program_outdated?(d, software)
			init_update(d)
		end
	when "registry_version"
		if registry_version_outdated?(d)
			init_update(d)
		end
	when "file_version"
		if file_version_outdated?(d)
			init_update(d)
		end
	else
		log_event(8, 'ERROR', "The update method specified ('#{d["method"]}') in the definition for '#{d["description"]}' is invalid.")
	end
end