#!/usr/bin/env ruby
require 'csv'
require 'win32ole' #require library for ActiveX Data Objects (ADO)


# Define and set global variable for database file
$database_file_path = '.\inventory.mdb' # For windows path


# CREATE CLASS AccessDb for database connection handling
class AccessDb
	# Set variables as accessors, so that they have read/writability
	attr_accessor :mdb, :connection, :data, :fields

	# Constructor for class AccessDb
	def initialize (mdb = nil)
		@mdb = mdb
		@connection = nil
		@data = nil
		@fields = nil
	end

	# Open the connection to Database
	def open
		connection_string = 'Provider=Microsoft.Jet.OLEDB.4.0;Data Source='
		connection_string << @mdb
		@connection = WIN32OLE.new('ADODB.Connection')
		@connection.Open(connection_string)
	end

	# Method for querying the Database
	def query(sql)
		recordset = WIN32OLE.new('ADODB.Recordset')
		recordset.Open(sql, @connection)
		@fields = []
		recordset.Fields.each do |field|
			@fields << field.Name
		end

		begin
			# Transpose to have array of rows
			@data = recordset.GetRows.transpose
		rescue
			@data = []
		end

		recordset.Close
	end

	def get_length
		return data.length
	end

	# Method for executing a sql command
	def execute(sql)
		@connection.Execute(sql)
	end

	# Destructor method for AccessDb Class
	def close
		@connection.Close
	end
end


# LOAD DATABASE FUNCTION loads the database into variable connection
def load_database
	user_input = ""

	# Check if database file is located in cwd
	until (user_input == "Y" || user_input == "N")
		print "Is the database file in the current working directory and named 'inventory.mdb'? [Y/N]: "
		user_input = gets.strip.upcase
	end

	# If database file is located elsewhere get file path from user
	unless (user_input == "Y")
		puts "Please specify the pathname where the database file is located, including file name."
		puts "Reminder: this program will only work on .mdb databases, it will not work for .accdb."
		puts '(e.g. ..\inventory.mdb or c:\Documents\Inventory\inventory.mdb):'
		$database_file_path = gets.strip
	end

	begin
		# Attempt to connect to the database, and store that connection in db
		db = AccessDb.new($database_file_path)
		db.open
	rescue
		# Abort if database not found
		abort "Unable to continue - database file #{$database_file_path} not found."
	end

	# Return database connection handle
	return db
end


# HELP SCREEN is called when user puts ?|-h|help as a guide
def help_scrn
	# define the help variable which holds the help text
	help = 
	"Usage: ruby inventory.rb [?|-h|help|[-u|-o|-z <infile>|[<outfile>]]]\n
	Parameters:
	   ?                 displays this usage information
	   -h                displays this usage information
	   help              displays this usage information
	   -u <infile>       update the inventory using the file <infile>.
	                     The filename <infile> must have a .csv
	                     extension and it must be a text file in comma
	                     separated value (CSV) format. Note that the
	                     values must be in double quote.
	   -z|-o [<outfile>] output either the entire content of the
	                     database (-o) or only those records for which
	                     the quantity is zero (-z). If no <outfile> is
	                     specified then output on the console otherwise
	                     output in the text file named <outfile>. The
	                     output in both cases must be in a tab separated
	                     value (tsv) format."

	# print the help screen
	puts help
end


# UPDATE INVENTORY FUNCTION called when user puts -u <infile>
def update_inv
	# Strip .csv file from input if present
	if (ARGV[0] != nil)
		update_file = ARGV.shift

		# Ensure update file is .csv file
		if (!update_file.end_with?(".csv"))
			abort "Invalid file format -- unable to proceed."
		else
			# Attempt to open user csv file. If not found, abort program.
			begin
				csv_file = CSV.open(update_file, "r")
			# Instead of asking for new file name, abort if file not found.
			rescue
				abort "Input file #{update_file} not found - aborting."
			end
		end
	else
		abort "You must provide an input file with -u!"
	end

	# Create database connection handle
	db = load_database

	# Set variable to determine how many successful updates there were
	success_count = 0

	# Go through db_values and update Database
	CSV.foreach(update_file) do |a|
		begin
			# Generate query string, field names, and rows of data from database
			execute_string = "INSERT INTO items(barcode,item_name,item_category,quantity,price,description) VALUES('#{a[0]}', '#{a[1]}', '#{a[2]}', #{a[3]}, #{a[4]}, '#{a[5]}');"
			# Attempt to execute the query string
			db.execute(execute_string)
			success_count += 1
		rescue
			next
		end
	end
	failed_count = csv_file.count - success_count

	# Get correct pluralization
	if (success_count == 1)
		plural_succeed = "record"
	else
		plural_succeed = "records"
	end

	if (failed_count == 1)
		plural_fail = "was 1 item"
	else
		plural_fail = "were #{failed_count} items"
	end
	
	# Create success message
	success_msg = "\nUpdated #{success_count} database #{plural_succeed} successfully."

	# Add message if there were problems updating
	unless (failed_count == 0)
		success_msg << "\nThere #{plural_fail} that could not be uploaded to the database.\n\nEither #{update_file} is not formatted propery, or there were duplicates in the database."
		success_msg << "\nConsider updating each individual record quantity if needed."
	end
	# Display update information to user
	puts success_msg
	# Close database connection
	db.close
end

def format_output
	puts "+----------------------+------------------------------------------+----------------------+----------+---------+-------------------------------------------+"
	puts "| Barcode".ljust(23) 			<< "| Item Name:".ljust(43) <<
		 "| Item Category".ljust(23) 	<< "| Quantity".ljust(11) 	<<
		 "| Price".ljust(10)			<< "| Description".ljust(43) << " |"
	puts "+----------------------+------------------------------------------+----------------------+----------+---------+-------------------------------------------+"
end


# OUTPUT INVENTORY FILE called when user puts -o|-z <outfile>
def load_file(view_option)
	# Strip .tsv file from input if present
	if (ARGV[0] != nil)
		output_file = ARGV.shift
		# Ensure outfile is of type .tsv
		unless (output_file.end_with?(".tsv"))
			abort "<outfile> file format must be .tsv!"
		end		
	end

	# Create database connection handle
	db = load_database

	# Generate query string
	query_string = "SELECT * FROM items"

	# Append query string to check only for quantity = 0
	# if -z was given (view_option == false)
	if (view_option == false)
		query_string << " WHERE quantity = 0"
	end

	# Run the query and store the data
	db.query(query_string)
	rows = db.data

	# If no rows returned (typically only when specified with -z but none exist)
	if (db.get_length < 1)
		puts "No database records found."
	else
		# If user didn't enter outfile, print to screen
		if (output_file == nil)
			format_output
			# Loop through the rows and output data
			rows.each do |a|
				print "| #{a[0]}".ljust(23)
				print "| #{a[1]}".ljust(43)
				print "| #{a[2]}".ljust(23)
				print "| #{a[3]}".ljust(11)
				print "| #{a[4]}".ljust(10)
				print "| #{a[5]}".ljust(43) << " |"
				print "\n"
			end
			puts "+----------------------+------------------------------------------+----------------------+----------+---------+-------------------------------------------+"
		# If user specified an outfile, print to outfile
		else
			# Write to file by first opening file and writing over it
			CSV.open(output_file, "w", {:col_sep => "\t"}) do |csv|
				rows.each do |a|
				  csv << [a[0], a[1], a[2], a[3], a[4], a[5]]
				end
			end
			puts "File was successfully updated!"
		# If user specified file was not .tsv
		end
	end
	# Close database connection
	db.close
end


# SEARCH INVENTORY FILE called when user enters "ruby inventory.rb" and gets barcode
def search_inv(barcode,db)

	# Initialize variable for user input
	user_input = ""
	# Generate query string
	query_string = "SELECT * FROM items WHERE barcode = '#{barcode}'"

	# Run the query and store the data
	db.query(query_string)
	field_names = db.fields
	rows = db.data

	# If item was found in database
	if (rows.length > 0)
		# If the quantity of found barcode is zero, give user option to update quantity
		if (rows[3] == 0)
			print = "\nBarcode #{barcode} found in the database but has a zero quantity. Do you want to update quantity? [Y/N]: "

			# Validate user input
			until (user_input == "Y" || user_input == "N")
				user_input = gets.strip.upcase
			end

			# If user chooses yes, call update_db_entry function
			if (user_input == "Y")
				update_db_entry
			end
			# Else continue with printing information
		else
			puts "\nBarcode #{barcode} found in the database. Details are given below:"
			puts "   Item Name: #{rows[0][1]}"
			puts "   Item Category: #{rows[0][2]}"
			puts "   Quantity: #{rows[0][3]}"
			puts "   Price: #{rows[0][4]}"
			puts "   Description: #{rows[0][5]}"
		end

	else
		puts "No items"
	end

	# Close database connection
	db.close
end


# Main loop
user_option = ARGV.shift

if (user_option == '?' || user_option == '-h' || user_option == 'help')
	help_scrn
elsif (user_option == '-u')
	# -u <infile> updates the inventory (LINE 126)
	update_inv
elsif (user_option == '-z' || user_option == '-o')
	if (user_option == '-z')
		load_file(false)
	else
		load_file(true)
	end

# If user scans barcode or enters one manually into the command line
else
	if (user_option != nil)
		input = user_option.strip
	else
		print "Barcode number: "
		input = gets.strip
	end

	# Test if barcode is sequence of numbers
	if (input =~ /^[0-9]{1,20}$/)
		db_conn = load_database
		search_inv(input,db_conn)
	else
		abort "Usage: ruby inventory.rb [?|-h|help|[-u|-o|-z <infile>|[<outfile>]]]\nBarcode must be only digits 0-9 of maximum length 20."
	end
end