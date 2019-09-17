# module SheetExtraction
#     extend ActiveSupport::Concern
#     extend ProgramTableLocation
#   # @xlsx is already obtained in read_sheet. 

#   def get_program_table( bank_name )
#   	# Check if the table still stays the same as before
#   	# if check_table_location() # K needs to think about this.

#       # Loop over all the tabs in the bank rate sheet 
#       sheet_hash = SheetExtraction.table_location[ bank_name ]
#       sheet_hash.keys.each do |sheet_name|  # e.g., key = 'List_of_rrcc_NewRez_CFR'
#         bank_hash[ sheet_name ].each do | rrcc_list |
#         #SheetExtraction.table_location.values.each do | rrcc_list |
#           extract_program_table( sheet_name, rrcc_list )
#         end
#       end

#       # List_of_row_col_NewRez_CFR.each do | rrcc | # for rrcc in List_of_rrcc_NewRez_CFR
#       #     extract_program_table( @sheet_name, rrcc )
#       # end
#   	# else
#   	# 	Raise warning: “Table format has been changed in the sheet, %s, of the new bank file!”, @sheet_name. # Need manual correction for the table locations in program_tables_location.rb
#   end
#   # def process_adjustment_table(  )
#   # 	# Put those adj tables that are universal in at least most of the banks here. For other not so commonly used tables, we have to process them in each tab/sheet function respectively. 
#   # 	# K will fill this out later. 
#   # end
 
#   # def func_col_key_program_table( raw_col_key )
#   #   # if raw_col_key[ 1 ] = " "
#   #   #   return raw_col_key[ 0 ] # e.g. Term = 5 yr
#   #   # elsif raw_col_key[ 2 ] = " "
#   #   #   return raw_col_key[ 0:1 ]
#   #   return raw_col_key[0:1]

#     # For cases where the lock period is 30-digit, we will make change to this function later. 

#     # if raw_col_key.include?( " " )
#     #    return strsep( raw_col_key, "-" )[ 0 ]
#     # elsif raw_col_key.include?( "-" )

#   def extract_program_table( sheet_name, rrcc_list, func_col_key := lambda x: x[0:1] ) 
    
#     sheet_data = @xlsx.sheet( sheet_name )
    
#     # Get program table title
#     row_start = rrcc_list[ 0 ]
#     @title = sheet_data.cell(row_start,col_start)
#     @program = @sheet_obj.programs.new(program_name: @title)
#     p_name = @title + " " + sheet
#     @program.update_fields( p_name )
#     program_property( @title )	

#     # Get program table data
#     rrcc_list[ 0 ] = rrcc_list[ 0 ] + 1
#     extract_table_2D( sheet_name, rrcc_list, func_col_key := func_col_key )
    

#       #(row_start..row_end+1).each do |r|		
#         #     begin
#         #         row = sheet_data.row(r)

#         #         if r == row_start
#         #             @title = sheet_data.cell(r,col_start)
#         #             @program = @sheet_obj.programs.new(program_name: @title)
#         #             p_name = @title + " " + sheet
#         #             @program.update_fields( p_name )
#         #             program_property( @title )	
#         #         elsif r == row_start + 1
#         #             # Check if program table location is the same as before
#         #             cell_15d = sheet_data.cell( r, col_start + 1 )
#         #             if cell_15d == “15 Days” or cell_15d == “15 days”:			
#         #                 continue;
#         #             else
#         #                 Raise warning: Potential program table format change!
#         #             end
#         #         else
#         #           @block_hash = {}
#         #           interest_rate = ''
#         #           (r-row_start..row_end-row_start).each do |r|
#         #             @data = []
#         #             (0..(col_end - col_start)).each_with_index do |index, c_i|
#         #                 col = col_start + c_i
#         #                 value = sheet_data.cell(r,col)
#         #                 if ( c_i == 0 )
#         #                     if value.present?  
#         #                         interest_rate = value
#         #                         @block_hash[ interest_rate ] = {}
#         #                     else
#         #                         break # skip the rest of the columns and move to the next row. Double-check!!!
#         #                     end
#         #                 else
#         #                     @block_hash[ interest_rate ][ 15 * c_i ] = value
#         #                 end
                    
#         #                 @data << value
#         #             end
#         #           end

#         #           if @data.compact.length > 0 and r == row_end + 1
#         #               Raise warning: Potential program table row number change and more data from the table to process! (Need to modify row_end)
#         #           end

#         #           # if @block_hash.values.first.keys.first.nil?
#         #           #     @block_hash.values.first.shift
#         #           # end

#         #           @block_hash.delete(nil) # We still need to show nil if the base point is nil in the rate sheet. Think about this!!!
#         #           error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: col_start, loan_category: @sheet_name, error_detail: e.message)
#         #           error_log.save
#         #         end
#         #     end
#       # end
#   # end

#     # Call save_rate_sheet_table( params, program_or_adj == "Program" )
# end

# def extract_adjustment_table_1D( sheet_name, row_start, row_end, col_start, col_end, func_row_key, func_col_key )



# def extract_table_2D( sheet_name, rrcc_list, func_row_key := lambda x : x, func_col_key := lambda x : x, use_ext_col_keys := 0, row_for_col_key_list := [], col_key_list := [] ) 
#   # rrcc_list = [ row_start, row_end, col_start, col_end ]
#   # use_ext_col_keys: 0: Not use. 1: use row that contains the keys. 2: use list of keys, which starts with index 0. 

#   row_start = rrcc_list[ 0 ]
#   row_end   = rrcc_list[ 1 ]
#   col_start = rrcc_list[ 2 ]
#   col_end   = rrcc_list[ 3 ]
  
#   # @xlsx.sheets.each do |sheet|
#   #   if ( sheet == sheet_name )
#   #       @sheet_name = sheet
#   #       sheet_data = @xlsx.sheet(sheet)
#   #       break;
#   #   end
#   # end
  
#   sheet_data = @xlsx.sheet( sheet_name )

#   # Add some check routines here for whether the table has been changed. 
#   # To be added ???

#   (row_start..row_end+1).each do |r|		
#         begin
#             if r == row_start
#               if use_ext_col_keys <= 0
#                 @col_key_list = sheet_data.row(r)
#               else
#                 if use_ext_col_keys == 1              
#                   @col_key_list = row_for_col_key_list
#                 elseif use_ext_col_keys == 2 
#                   @col_key_list = col_key_list
#                 else
#                   @col_key_list = []
#                 end
#               end
#               (col_start..col_end).each do |c|
#                 if sheet_data.cell(r,c).present? and c > col_start
#                   col_start2 = c
#                 end
#             else
#               @block_hash = {}
#               row_key = ''
#               (0..row_end-row_start).each do |r|
#                 @data = []
#                 (col_start..col_end).each do |c|
#                     #col = col_start + c_i
#                     value = sheet_data.cell(r,c)
#                     if ( c == col_start )
#                       if value.present?  
#                         row_key = func_row_key( value )
#                         @block_hash[ row_key ] = {}
#                       elsif 
#                           break # skip the rest of the columns and move to the next row. Double-check!!!
#                       end
#                     else
#                       if c >= col_start2
#                         if use_ext_col_keys == 2 
#                           ck = c - col_start2
#                         else
#                           ck = c
#                         end
#                         @block_hash[ row_key ][ func_col_key( @col_key_list[ c ] ) ] = value
#                       end
#                     end
#                     @data << value
#                 end
#               end

#               if @data.compact.length > 0 and r == row_end + 1
#                   Raise warning: Potential table row number change and more data from the table to process! (Need to modify row_end)
#               end

#               # if @block_hash.values.first.keys.first.nil?
#               #     @block_hash.values.first.shift
#               # end

#               @block_hash.delete(nil) # We still need to show nil if the base point is nil in the rate sheet. Think about this!!!
#               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: col_start, loan_category: @sheet_name, error_detail: e.message)
#               error_log.save
#             end
#         end
#   end
# end

# # Call save_rate_sheet_table( params, program_or_adj == "Program" )
# end


# def get_adj_LTV_FICO_table( sheet_name, row_start, row_end, col_start, col_end, func_row_key, func_col_key ) 
#   extract_table_2D( sheet_name, rrcc_list, func_row_key := func_row_key, func_col_key := func_col_key )
