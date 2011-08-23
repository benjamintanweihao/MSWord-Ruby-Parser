#!/usr/bin/ruby 
# Takes a preprocessed HTML document, splits the document into separate questions,
# then re-assembles then into individual HTML files
# Note: This version only works on Windows, sadly.
# @author: Benjamin Tan

require 'hpricot'
require 'open-uri'
require 'tidy'
Tidy.path = "C:\\Ruby\\lib\\ruby\\gems\\1.8\\gems\\tidy-1.1.2\\lib\\tidy\\tidy.dll"

FILE_TO_PARSE = "./test.htm"
TEMPLATE_FILE = "./math_template/template.htm"
START_DELIMITER = "###"
BODY_DELIMITER = "## BODY ##"
DIRECTORY_DELIMITER = "MPBodyInit(## DIRECTORY ##)"

class Parser	

	def self.parse_questions( questions_file )
		nice_html = ""
		ugly_html = ""
		question_html = ""
		questions_html_array = Array.new
		
		doc = Hpricot(open( questions_file ))
		doc.search("div[@class*=Section1]").each do |line|
			ugly_html << line.inner_html
		end
		
		Tidy.open do |tidy|
			tidy.options.wrap = 0
			tidy.options.indent = 'auto'
			tidy.options.indent_attributes = false
			tidy.options.indent_spaces = 4
			tidy.options.vertical_space = false
			tidy.options.char_encoding = 'utf8'
			tidy.options.bare = true
			tidy.options.clean = true
			nice_html = tidy.clean( ugly_html )
		end
		
		nice_html = nice_html.strip.gsub(/\n+/, "\n")
		nice_html.each do |line|
			if line.include? START_DELIMITER
				questions_html_array << question_html
				question_html = ""
			else
				question_html << line
			end
		end
		questions_html_array
	end
	
	def self.get_template_html( file_path )
		file_contents = ""
		File.open( file_path ).each do |line|
			file_contents << line
		end
		file_contents
	end
	
	def self.get_assembled_question( template_html, images_directory, question_html )
		assembled_question = ""
	
		template_html.each do |line|
			if line.include? DIRECTORY_DELIMITER
				assembled_question << "MPBodyInit('"+images_directory+"')"
			elsif line.include? BODY_DELIMITER
				assembled_question << question_html	
			else
				assembled_question << line
			end	
		end
		assembled_question
	end
	
	def self.write_question_to_file( file_name, question)
		File.open( "./Processed/"+file_name+".htm", 'w') {|f| f.write( question )}
	end
	
	def self.run
		questions_array = Parser.parse_questions( FILE_TO_PARSE )
		template_html = Parser.get_template_html( TEMPLATE_FILE )
		
		questions_array.each_with_index do |question, index|
			assembled_question = Parser.get_assembled_question( template_html, './Maths Mid term test_files', question )
			Parser.write_question_to_file( "Maths_Mid_Term_Qn_"+index.to_s, assembled_question);
		end
	end
end


