
require 'xml'
require 'zip'
require 'highline/import'


def mac_compress
	exec('ditto -ck --rsrc --sequesterRsrc template file.pptx')
end

def edit_text

	target = ask "Targeted text: "
	substitute= ask "Substituted text: "

	text = File.read('template/ppt/slides/slide1.xml').gsub(target,substitute)

	if text!=""
		puts text
	#	File.open("template/ppt/slides/slide1.xml", "w").write(text)
		File.open("template/ppt/slides/slide1.xml", 'w') { |file| file.write(text) }

	else
		puts 'Nothing is found'
	end
 
end


edit_text
mac_compress

