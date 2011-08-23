# Takes in a raw MS Word document, apply delimters, 
# then converts the file into HTML, convert equations into gifs.
# @author Benjamin Tan
require 'win32ole'

class Preprocessor
	
	def self.setup( file_path )
		$wsh = WIN32OLE.new('Wscript.Shell')
		$word = WIN32OLE.new('Word.Application')
		$word.Visible = true
		$word.DisplayAlerts = false;
		fso = WIN32OLE.new("Scripting.FileSystemObject")
		path = fso.GetAbsolutePathName( file_path )
		doc = $word.Documents.Open(path)
	end
	
	def self.run_macros
		macros_to_run_array = Array.new
		macros_to_run_array << "Split_Questions_Num_Bracket"
		macros_to_run_array << "Split_Questions_Num_Letter"
		
		macros_to_run_array.each_with_index do |macro_to_run, index|
			puts "Running Macro: " + macro_to_run
			$word.run( macro_to_run )
			sleep(1)
		end
	end
	
	def self.select_publish_as_math_page
		sleep(3)
		puts $wsh.AppActivate('Word')
		# Navigate to menu and select "Publish as Math Page"
		sleep(2)
		$wsh.SendKeys('%')
		$wsh.SendKeys('F') 
		sleep(2) 
		$wsh.SendKeys('U') 
		sleep(2) 
		$wsh.SendKeys('{UP}')
		sleep(2)
		$wsh.SendKeys('{ENTER}')
		sleep(3)
		$wsh.SendKeys('{ENTER}')
	end
	
	def self.run( file_path )
		Preprocessor.setup( file_path )
		Preprocessor.run_macros
		sleep(2)
		Preprocessor.select_publish_as_math_page
	end
	#doc.Close
	#$word.Quit
end

