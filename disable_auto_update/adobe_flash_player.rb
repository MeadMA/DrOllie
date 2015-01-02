def mod_mms(filename)
	mms = Hash.new
	if File.exist?(filename)
		content = File.read(filename)
		content = content.gsub("\r\n", "\n")
		content = content.split("\n")
		content.each do |line|
			item = line.split('=')
			mms[item[0]] = item[1]
		end
	end
	data = ""
	mms['AutoUpdateDisable'] = '1'
	mms['SilentAutoUpdateEnable'] = '0'
	mms.each do |setting, value|
		data << setting + "=" + value + "\r\n"
	end
	if File.write(filename, data)
		return true
	else
		return false
end

# If mms.cfg files exist, modify them to disable automatic updates
# If the files don't exist, create them and disable automatic updates
mod_mms(ENV['SystemRoot'].gsub("\\", "/") + "/System32/macromed/flash/mms.cfg")
if Config::CONFIG['host_cpu'] == 'x86_64'
	mod_mms(ENV['SystemRoot'].gsub("\\", "/") + "/SysWOW64/macromed/flash/mms.cfg")
end

# Delete scheduled task
`schtasks /delete /tn "Adobe Flash Player Updater" /f`

# Delete service
`sc delete AdobeFlashPlayerUpdateSvc`