# Takes in a raw MS Word document, apply delimters, 
# then converts the file into HTML, convert equations into gifs.
# @author Benjamin Tan
require 'win32ole'

$wsh = WIN32OLE.new('Wscript.Shell')
$word = WIN32OLE.new('Word.Application')

$word.Visible = true
$word.DisplayAlerts = false;
fso = WIN32OLE.new("Scripting.FileSystemObject")
path = fso.GetAbsolutePathName("testdoc.doc")
doc = $word.Documents.Open(path)


def run_macros
	macros_to_run_array = Array.new
	macros_to_run_array << "Split_Questions_Num_Bracket"
	macros_to_run_array << "Split_Questions_Num_Letter"
	
	macros_to_run_array.each_with_index do |macro_to_run, index|
		puts "Running Macro: " + macro_to_run
		$word.run( macro_to_run )
	end
end

def select_publish_as_math_page
	$wsh.AppActivate( 'Word')
	# Navigate to menu and select "Publish as Math Page"
	sleep (5)
	$wsh.SendKeys('%')
	$wsh.SendKeys('F') 
	sleep (2) 
	8.times { $wsh.SendKeys('{DOWN}')}	
	$wsh.SendKeys('{RIGHT}')
	3.times { $wsh.SendKeys('{DOWN}')}
	$wsh.SendKeys('{ENTER}')
	sleep(1)
	$wsh.SendKeys('{ENTER}')  
end

run_macros
select_publish_as_math_page

#doc.Close
#$word.Quit
