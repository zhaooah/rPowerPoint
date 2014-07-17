
require 'xml'
require 'zip'
require 'highline/import'

#Pares xml
#source = XML::Parser.file('template/ppt/slides/slide1.xml')
#content = source.parse
#target = content.find('hao')


def get_dir
	#Get current directory
	puts File.join(File.absolute_path("/")  , 'template')


end


def package
	directory = '/template/'
	filename = 'new.zip'

	Zip::File.open(filename, Zip::File::CREATE) do |zipfile|
	    Dir[File.join(directory, '**', '**')].each do |file|
	      zipfile.add(file.sub(directory, ''), file)
	    end
	end

end


def edit_text
	
	file = File.open('template/ppt/slides/slide1.xml', "w+")
	data = file.read
	
	puts data

	target = ask "Targeted text: "
	substitute= ask "Substituted text: "


	replace=data.gsub(target,substitute)

	file.puts replace

	file.close
end

def new_edit_text

	target = ask "Targeted text: "
	substitute= ask "Substituted text: "

	text = File.read('template/ppt/slides/slide1.xml').gsub(target,substitute)
	File.open("template/ppt/slides/slide1.xml", "w").write(text)
end

def rzip
	  Zip::Archive.open('filename.zip', Zip::CREATE) do |ar|
      ar.add_dir('dir')
    
      Dir.glob('dir/**/*').each do |path|
        if File.directory?(path)
          ar.add_dir(path)
        else
          ar.add_file(path, path) # add_file(<entry name>, <source path>)
        end
      end
    end

end

new_edit_text


#package


