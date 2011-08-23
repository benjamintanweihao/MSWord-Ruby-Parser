require 'ftools'
require 'preprocessor.rb'
require 'parser.rb'

ORIGINAL_PATH = "./Original/"
PROCESSED_PATH = "./Processed/"
PROCESSED_FILENAME = "processed"
UNPROCESSED_PATH = "./Unprocessed/"

# Fix image links after file has been processed
def fix_image_links( file_name )
	lines = File.readlines(file_name)
    lines.each do |line| 
    	line.gsub!(/processed_files/, File.basename( file_name, '.htm') + "_files")
	end
	File.open( file_name, 'w') { |f| f.write lines }
end

def run
	files_to_process = Array.new
	
	puts "Searching for *UNPROCESSED* MS Word files ..."
	Dir[ UNPROCESSED_PATH + "*.doc*" ].each do |file|
		if File.exists? ORIGINAL_PATH + File.basename(file)
			puts "\t[exists,skipping]: '"+file+"' "
		else
			puts "\t[added to queue ]: '"+file+"' "
			files_to_process << File.expand_path(file)
		end
	end
	
	files_to_process.each do |f|
		# Run the preprocessor
		Preprocessor.run( f )
		# Rename the processed files, since MathType is not smart enough to rename them
		processed_file_name = PROCESSED_PATH + PROCESSED_FILENAME + ".htm"
		
		while File.size?( File.expand_path( processed_file_name ) )==0 or File.size?( File.expand_path(processed_file_name) )==nil
			puts "\tStill waiting for processed file"
			sleep(5)
		end
		
		renamed_path = PROCESSED_PATH + File.basename( f, '.*')
		File.rename( File.expand_path( PROCESSED_PATH + PROCESSED_FILENAME + ".htm" ), 
					 File.expand_path( renamed_path + ".htm" ) )
		File.rename( File.expand_path( PROCESSED_PATH + PROCESSED_FILENAME + "_files" ), 
					 File.expand_path( renamed_path + "_files" ) )		
		
		# Copy file to 'Original' folder
		File.copy( f, ORIGINAL_PATH + File.basename(f) )
		
		# Remove un-needed 'processed.htm and 'processed_files' folder
		if File.exist?(  PROCESSED_PATH + PROCESSED_FILENAME + ".htm" )
			puts "Removing temporary files"
			File.delete(File.expand_path( PROCESSED_PATH + PROCESSED_FILENAME + ".htm") )
			File.delete(File.expand_path( PROCESSED_PATH + PROCESSED_FILENAME + "_files") )
		end
		# Fix relative paths for images
		fix_image_links( File.expand_path( renamed_path + ".htm" ) ) 
	end
end

run