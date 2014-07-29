
#require 'highline/import'
require 'csv'
require 'json'



#The linux 'zip' command is not working on Mac OS X.
def renmae_zip
	exec('mv template.pptx template.zip')
end

def mac_unzip
	exec('ditto -x -k template.zip template')
end


def file_write(text,filename)
	if text!=""
		#puts text
		File.open(filename, 'w') { |file| file.write(text) }
	else
		puts 'Error!Cannot find placeholder'
	end
end


#Copy slides{num}.xml to ppt/slides
def copy_slides(dest,source)

	text = File.read('template/ppt/slides/slide'+source.to_s+'.xml')
	file_write(text,'template/ppt/slides/slide'+dest.to_s+'.xml')


	text = File.read('template/ppt/slides/_rels/slide'+source.to_s+'.xml.rels')
	file_write(text,'template/ppt/slides/_rels/slide'+dest.to_s+'.xml.rels')

	return dest
end

#[Content_Types].xml
def edit_content_type(slide_number)

	content_type_xml='<Override PartName="/ppt/slides/slide'+slide_number.to_s+'.xml" ContentType="application/vnd.openxmlformats-officedocument.presentationml.slide+xml"/></Types>'

	text = File.read('template/[Content_Types].xml').gsub('</Types>',content_type_xml)
	file_write(text,"template/[Content_Types].xml")

end


#presentation.xml.rels
#This document contains information about the relations between the document part and all the subsidiary parts, and the code snippet uses this information to find each of the slides, so that it can retrieve the title from the slide and compare it to the title you supply. 
def set_document_rels(slide_number)

	rid=1000+slide_number
	new_rels_xml='<Relationship Id="rId'+rid.to_s+'" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/slide" Target="slides/slide'+slide_number.to_s+'.xml"/>
</Relationships>'
	text = File.read('template/ppt/_rels/presentation.xml.rels').gsub('</Relationships>',new_rels_xml)
	file_write(text,"template/ppt/_rels/presentation.xml.rels")

	return rid
end

#ppt/presentation.xml
def add_to_main_presentation(rid,slide_number)

	new_slidID_xml='<p:sldId id="'+(999+slide_number).to_s + '" r:id="rId'+rid.to_s + '"/> </p:sldIdLst>'
	text = File.read('template/ppt/presentation.xml').gsub('</p:sldIdLst>',new_slidID_xml)

	file_write(text,"template/ppt/presentation.xml")
end
  


def edit_text(placeholder,replacement,slide_number)

	text = File.read('template/ppt/slides/slide'+slide_number.to_s+'.xml').gsub(placeholder,replacement)
 	file_write(text,'template/ppt/slides/slide'+slide_number.to_s+'.xml')

end

def mac_compress
	exec('ditto -ck --rsrc --sequesterRsrc template file.pptx')
end

