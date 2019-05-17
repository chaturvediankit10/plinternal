# def du_refi_plus_fixed_rate
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus Fixed Rate")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       sub_data = ''
  #       primary_key = ''
  #       secondry_key = ''
  #       fixed_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       misc_key = ''
  #       adj_key = ''
  #       term_key = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed High Balance"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # High Balance
  #               jumbo_high_balance = false
  #               if @title.include?("High Balance")
  #                 jumbo_high_balance = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming, arm_basic: arm_basic, loan_category: sheet, jumbo_high_balance: jumbo_high_balance)
  #               
  #               @block_hash = {}
  #               key = ''
  #               main_key = ''
  #               if @program.term.present?
  #                 main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               else
  #                 main_key = "InterestRate/LockPeriod"
  #               end
  #               @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[main_key][key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[main_key][key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (63..94).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(75)
  #         if row.compact.count >= 1
  #           (3..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if (r == 65 && cc == 3)
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Loan Size Adjustments"
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All Fixed Confoming Adjustment
  #                 if r >= 66 && r <= 73 && cc == 8
  #                   fixed_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
  #                 end
  #                 if r >= 66 && r <= 73 && cc >8 && cc <= 19
  #                   fixed_data = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
  #                 end

  #                 # Subordinate Financing Adjustment
  #                 if r >= 76 && r <= 80 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 80 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 80 && cc > 6 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 83 && r <= 89 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 83 && r <= 89 && cc > 6 && cc <= 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 92 && r <= 94 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 92 && r <= 94 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   misc_key = value
  #                   @adjustment_hash[misc_key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 75 && r <= 83 && cc == 15
  #                   if value.include?("Condo")
  #                     adj_key = "Condo/75/15"
  #                   else
  #                     adj_key = value
  #                   end
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 75 && r <= 83 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r == 85 && cc == 13
  #                   adj_key = value
  #                   @adjustment_hash[adj_key] = {}
  #                 end
  #                 if r == 85 && cc == 17
  #                   @adjustment_hash[adj_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 89 && r <= 93 && cc == 16
  #                   adj_key = value
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 89 && r <= 93 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end


   # def du_refi_plus_fixed_rate_105
  #   @program_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus Fixed Rate_105")
    # sheet_data = @xlsx.sheet(sheet)
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("DU Refi Plus 10yr Fixed >125 LTV"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               
  #               @block_hash = {}
  #               key = ''
  #               main_key = ''
  #               if @program.term.present?
  #                 main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               else
  #                 main_key = "InterestRate/LockPeriod"
  #               end
  #               @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[main_key][key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[main_key][key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end

  #       #For Adjustments
  #       @xlsx.sheet(sheet).each_with_index do |sheet_row, index|
  #         index = index+ 1
  #         if sheet_row.include?("Loan Level Price Adjustments: See Adjustment Caps")
  #           (index..@xlsx.sheet(sheet).last_row).each do |adj_row|
  #             # First Adjustment
  #             if adj_row == 65
  #               key = ''
  #               key_array = []
  #               rr = adj_row
  #               cc = 3
  #               @occupancy_hash = {}
  #               main_key = "All Occupancies"
  #               @occupancy_hash[main_key] = {}

  #               (0..2).each do |max_row|
  #                 column_count = 0
  #                 rrr = rr + max_row
  #                 row = @xlsx.sheet(sheet).row(rrr)

  #                 if rrr == rr
  #                   row.compact.each do |row_val|
  #                     val = row_val.split
  #                     if val.include?("<")
  #                       key_array << 0
  #                     else
  #                       key_array << row_val.split("-")[0].to_i.round if row_val.include?("-")
  #                       key_array << row_val.split[1].to_i.round if row_val.include?(">")
  #                     end
  #                   end
  #                 end

  #                 (0..16).each do |max_column|
  #                   ccc = cc + max_column
  #                   begin
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                     if row.include?("All Occupancies > 15 Yr Terms")
  #                       if value != nil && value.to_s.include?(">") && value != "All Occupancies > 15 Yr Terms" && !value.is_a?(Numeric)
  #                         key = value.gsub(/[^0-9A-Za-z]/, '')
  #                         @occupancy_hash[main_key][key] = {}
  #                       elsif (value != nil) && !value.is_a?(String)
  #                         @occupancy_hash[main_key][key][key_array[column_count]] = value
  #                         column_count = column_count + 1
  #                       end
  #                     end
  #                   rescue Exception => e
  #                     error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                     error_log.save
  #                   end
  #                 end
  #               end
  #               make_adjust(@occupancy_hash, @program_ids)
  #             end

  #             # Second Adjustment(Adjustment Caps)
  #             if adj_row == 86
  #               key_array = ""
  #               rr = adj_row
  #               cc = 16
  #               @adjustment_cap = {}
  #               main_key = "Adjustment Caps"
  #               @adjustment_cap[main_key] = {}
  #               key = ''

  #               (0..4).each do |max_row|
  #                 column_count = 1
  #                 rrr = rr + max_row
  #                 row = @xlsx.sheet(sheet).row(rrr)
  #                 if rrr == 86
  #                   key_array = row.compact
  #                 end

  #                 (0..3).each do |max_column|
  #                   ccc = cc + max_column
  #                   begin
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                     if ccc == 16
  #                       key = value if value != nil
  #                       @adjustment_cap[main_key][key] = {} if value != nil
  #                     else
  #                       if !key_array.include?(value)
  #                         @adjustment_cap[main_key][key][key_array[column_count]] = value if value != nil
  #                         column_count = column_count + 1 if value != nil
  #                       end
  #                     end
  #                   rescue Exception => e
  #                     error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #                     error_log.save
  #                   end
  #                 end
  #               end
  #               make_adjust(@adjustment_cap, @program_ids)
  #             end

  #             # Third Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Max YSP")
  #               rr = adj_row
  #               cc = 4
  #               begin
  #                 @max_ysp_hash = {}
  #                 main_key = "Max YSP"
  #                 @max_ysp_hash[main_key] = {}
  #                 row = @xlsx.sheet(sheet).row(rr)
  #                 @max_ysp_hash[main_key] = row.compact[5]
  #                 make_adjust(@max_ysp_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # Fourth Adjustment (Adjustments Applied after Cap)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Loan Size Adjustments")
  #               rr = adj_row
  #               cc = 6
  #               @loan_size = {}
  #               main_key = "Loan Size / Loan Type"
  #               @loan_size[main_key] = {}

  #               (0..6).each do |max_row|
  #                 @data = []
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)
  #                   if key.present?

  #                     if (key.include?("<"))
  #                       key = 0
  #                     elsif (key.include?("-"))
  #                       key = key.split("-").first.tr("^0-9", '')
  #                     else
  #                       key
  #                     end
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     raise "value is nil at row = #{rrr} and column = #{ccc}" unless value || key
  #                     @loan_size[main_key][key] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@loan_size, @program_ids)
  #             end

  #             # Fifth Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               @cando_hash = {}
  #               main_key = "PropertyType/LTV/Term"
  #               @cando_hash[main_key] = {}

  #               (0..6).each do |max_row|
  #                 @data = []
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("Condo")
  #                     val = key.split
  #                     key1 = "Condo"
  #                     key2 = val[1].gsub(/[^0-9A-Za-z]/, '')
  #                     key3 = val[3].gsub(/[^0-9A-Za-z]/, '').split("yr")[0]
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @cando_hash[main_key][key1] = {}
  #                     @cando_hash[main_key][key1][key2] = {}
  #                     @cando_hash[main_key][key1][key2][key3] = value
  #                   end

  #                   if key == "Manufactured Home"
  #                     key1 = "Manufactured Home"
  #                     key2 = 0
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @cando_hash[main_key][key1] = {}
  #                     @cando_hash[main_key][key1][key2] = {}
  #                     @cando_hash[main_key][key1][key2] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@cando_hash, @program_ids)
  #             end

  #             # Sixth Adjustment(Misc Adjusters (2-4 Units))
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #                 rr = adj_row
  #                 cc = 15
  #               begin
  #                 @unit_hash = {}
  #                 main_key = "PropertyType/LTV"
  #                 @unit_hash[main_key] = {}

  #                 rrr = rr + 1
  #                 ccc = cc
  #                 key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                 if key.include?("Units")
  #                   key1 = "2-4 unit"
  #                   value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                   @unit_hash[main_key][key1] = {}
  #                   @unit_hash[main_key][key1] = value
  #                 end
  #                 make_adjust(@unit_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end


  #             # Seventh Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               begin
  #                 @data_hash = {}
  #                 main_key = "MiscAdjuster"
  #                 @data_hash[main_key] = {}

  #                 (0..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if !key.include?("Units")
  #                     key1 = key.include?(">") ? key.split(" >")[0] : key
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @data_hash[main_key][key1] = {}
  #                     @data_hash[main_key][key1] = value
  #                   end
  #                 end
  #                 make_adjust(@data_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # LTV Adjustment(Misc Adjusters)
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Adjustments Applied after Cap")
  #               rr = adj_row
  #               cc = 15
  #               @ltv_hash = {}
  #               main_key = "LTV"
  #               @ltv_hash[main_key] = {}

  #               (0..6).each do |max_row|
  #                 rrr = rr + max_row
  #                 ccc = cc
  #                 begin
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("LTV") && !key.include?("Condo")
  #                     key1 = key.split[1].to_i.round
  #                     key2 = key.include?("<") ? 0 : 30
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+4)
  #                     @ltv_hash[main_key][key1] = {} if @ltv_hash[main_key] == {}
  #                     @ltv_hash[main_key][key1][key2] = {}
  #                     @ltv_hash[main_key][key1][key2] = value
  #                   end
  #                 rescue Exception => e
  #                   error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                   error_log.save
  #                 end
  #               end
  #               make_adjust(@ltv_hash, @program_ids)
  #             end

  #             # CA Escrow Waiver Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Expanded Approval **")
  #               rr = adj_row
  #               cc = 3
  #               begin
  #                 @misc_adjuster = {}
  #                 main_key = "MiscAdjuster"
  #                 @misc_adjuster[main_key] = {}

  #                 (0..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?("CA Escrow Waiver") || key.include?("Expanded Approval **")
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+7)
  #                     @misc_adjuster[main_key][key] = {}
  #                     @misc_adjuster[main_key][key] = value
  #                   end
  #                 end
  #                 make_adjust(@misc_adjuster, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rrr, column: ccc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end

  #             # Subordinate Financing Adjustment
  #             if @xlsx.sheet(sheet).row(adj_row).include?("Subordinate Financing")
  #               rr = adj_row
  #               cc = 6
  #               begin
  #                 @subordinate_hash = {}
  #                 main_key = "FinancingType/LTV/CLTV/FICO"
  #                 key1 = "Subordinate Financing"

  #                 sub_key1 = row.compact[2].include?("<") ? 0 : row.compact[2].split(" ")[1].to_i
  #                 sub_key2 = row.compact[3].include?(">") ? row.compact[3].split(" ")[1].to_i : row.compact[3].to_i

  #                 @subordinate_hash[main_key] = {}
  #                 @subordinate_hash[main_key][key1] = {}

  #                 (1..2).each do |max_row|
  #                   rrr = rr + max_row
  #                   ccc = cc
  #                   key = @xlsx.sheet(sheet).cell(rrr,ccc)

  #                   if key.include?(">") || key == "ALL"
  #                     key2 = (key.include?(">")) ? key.gsub(/[^0-9A-Za-z]/, '') : key
  #                     value = @xlsx.sheet(sheet).cell(rrr,ccc+3)
  #                     value1 = @xlsx.sheet(sheet).cell(rrr,ccc+4)

  #                     @subordinate_hash[main_key][key1][key2] ={}
  #                     @subordinate_hash[main_key][key1][key2][sub_key1] = value
  #                     @subordinate_hash[main_key][key1][key2][sub_key2] = value1
  #                   end
  #                 end
  #                 make_adjust(@subordinate_hash, @program_ids)
  #               rescue Exception => e
  #                 error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #                 error_log.save
  #               end
  #             end
  #           end
  #         end
  #       end
  #     end
  #   end
  #   create_program_association_with_adjustment(@sheet)
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

    # def du_refi_plus_arms
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "Du Refi Plus ARMs")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       fixed_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       sub_data = ''
  #       misc_key = ''
  #       adj_key = ''
  #       term_key = ''
  #       @sheet = sheet
  #       (1..35).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               arm_basic = false
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               # High Balance
  #               if @title.include?("High Balance")
  #                 jumbo_high_balance = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet, jumbo_high_balance: jumbo_high_balance)
  #               
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (37..70).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(39)
  #         @sub_data = sheet_data.row(49)
  #         if row.compact.count >= 1
  #           (3..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All DU Refi Plus Conforming ARMs (All Occupancies)" || value == "Subordinate Financing" || value == "Loan Size Adjustments"
  #                   secondry_key = value
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All du refi plus Adjustment
  #                 if r >= 40 && r <= 47 && cc == 8
  #                   fixed_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key] = {}
  #                 end
  #                 if r >= 40 && r <= 47 && cc >8 && cc <= 19
  #                   fixed_data = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][fixed_key][fixed_data] = value
  #                 end

  #                 # Subordinate Financing Adjustment
  #                 if r >= 50 && r <= 54 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 50 && r <= 54 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 50 && r <= 54 && cc > 6 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 56 && r <= 57 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 56 && r <= 57 && cc == 8
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 60 && r <= 66 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 60 && r <= 66 && cc > 6 && cc <= 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 69 && r <= 70 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 69 && r <= 70 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   misc_key = value
  #                   @adjustment_hash[misc_key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 49 && r <= 58 && cc == 15
  #                   if value.include?("Condo")
  #                     adj_key = "Condo/75"
  #                   else
  #                     adj_key = value
  #                   end
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 49 && r <= 58 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 62 && r <= 64 && cc == 16
  #                   adj_key = value
  #                   @adjustment_hash[misc_key][adj_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 62 && r <= 64 && cc == 19
  #                   @adjustment_hash[misc_key][adj_key][term_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @sheet)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

    # def lp_open_access
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Access")
    # sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       @unit_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       unit_key = ''
  #       caps_key = ''
  #       term_key = ''
  #       max_key = ''
  #       fixed_key = ''
  #       sub_data = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access Super Conforming 10 Yr Fixed"))
  #           rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # rate type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               arm_basic = false
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae =false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end

  #       # Adjustment
  #       (63..97).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(73)
  #         @unit_data = sheet_data.row(82)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)

  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Number Of Units"
  #                   secondry_key = "PropertyType/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == 'Loan Size Adjustments'
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All fixed Adjustment
  #                 if r >= 66 && r <= 71 && cc == 8
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 66 && r <= 71 && cc > 8 && cc <= 19 && cc != 15
  #                   fixed_key = @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Adjustment
  #                 if r >= 74 && r <= 80 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 74 && r <= 80 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 74 && r <= 80 && cc >= 9 && cc <= 10
  #                   fixed_key = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
  #                 end

  #                 # Number of unit Adjustment
  #                 if r >= 83 && r <= 84 && cc == 3
  #                   unit_key = value
  #                   @adjustment_hash[primary_key][secondry_key][unit_key] = {}
  #                 end
  #                 if r >= 83 && r <= 84 && cc > 3 && cc <= 7
  #                   fixed_key = get_value @unit_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][unit_key][fixed_key] = value
  #                 end

  #                 # Loan Size Adjustments
  #                 if r >= 87 && r <= 93 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 87 && r <= 93 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustment
  #                 if r >= 95 && r <= 97 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 95 && r <= 97 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   @key = value
  #                   @adjustment_hash[primary_key][@key] = {}
  #                 end

  #                 # Misc Adjustment
  #                 if r >= 73 && r <= 80 && cc == 15
  #                   if value.include?("Condo")
  #                     cltv_key = "Condo=>75.01=>15.01"
  #                   else
  #                     cltv_key = value
  #                   end
  #                   @adjustment_hash[primary_key][@key][cltv_key] = {}
  #                 end
  #                 if r >= 73 && r <= 80 && cc == 19
  #                   @adjustment_hash[primary_key][@key][cltv_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r > 86 && r <= 90 && cc == 16
  #                   caps_key = value
  #                   @adjustment_hash[primary_key][@key][caps_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
  #                 end
  #                 if r > 86 && r <= 90 && cc == 19
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
  #                 end


  #                 if r == 93 && cc == 12
  #                   max_key = value
  #                   @adjustment_hash[primary_key][max_key] = {}
  #                 end
  #                 if r == 93 && cc == 16
  #                   @adjustment_hash[primary_key][max_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end


    # def lp_open_access_105
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Access_105")
  #       sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       term_key = ''
  #       caps_key = ''
  #       max_key = ''
  #       fixed_key = ''
  #       @sheet = sheet
  #       (1..61).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("LP Open Access 10yr Fixed >125 LTV"))
  #           rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # interest type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # interest sub type
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #               # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustment
  #       (63..86).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(65)
  #         @sub_data = sheet_data.row(68)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All Fixed Conforming > 15yr Terms (All Occupancies)"
  #                   secondry_key = "LoanSize/LoanType/Term/FICO/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Subordinate Financing"
  #                   secondry_key = "FinancingType/LTV/CLTV/FICO"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == "Number Of Units"
  #                   secondry_key = "PropertyType/LTV"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end
  #                 if value == 'Loan Size Adjustments'
  #                   secondry_key = "Loan Size Adjustments"
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All Fixed Conforming Adjustments
  #                 if r == 66 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r == 66 && cc > 6 && cc <= 19 && cc != 15
  #                   fixed_key = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Financing
  #                 if r == 69 && cc == 5
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r == 69 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r == 69 && cc >= 9 && cc <= 10
  #                   fixed_key = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][fixed_key] = value
  #                 end

  #                 # Number Of Units
  #                 if r >= 72 && r <= 73 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 72 && r <= 73 && cc == 5
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 76 && r <= 82 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 76 && r <= 82 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r >= 84 && r <= 86 && cc == 3
  #                   ltv_key = value
  #                   @adjustment_hash[primary_key][ltv_key] = {}
  #                 end
  #                 if r >= 84 && r <= 86 && cc == 10
  #                   @adjustment_hash[primary_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             begin
  #               value = sheet_data.cell(r,cc)
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   @key = value
  #                   @adjustment_hash[primary_key][@key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 68 && r <= 72 && cc == 15
  #                   if value.include?("Condo")
  #                     cltv_key = "Condo=>105=>15.01"
  #                   else
  #                     cltv_key = value
  #                   end
  #                   @adjustment_hash[primary_key][@key][cltv_key] = {}
  #                 end
  #                 if r >= 68 && r <= 72 && cc == 19
  #                   @adjustment_hash[primary_key][@key][cltv_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r > 76 && r <= 79 && cc == 16
  #                   caps_key = value
  #                   @adjustment_hash[primary_key][@key][caps_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = {}
  #                 end
  #                 if r > 76 && r <= 79 && cc == 19
  #                   @adjustment_hash[primary_key][@key][caps_key][term_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r == 82 && cc == 12
  #                   max_key = value
  #                   @adjustment_hash[primary_key][max_key] = {}
  #                 end
  #                 if r == 82 && cc == 16
  #                   @adjustment_hash[primary_key][max_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @program_ids)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end

    # def lp_open_acces_arms
  #   @programs_ids = []
  #   @xlsx.sheets.each do |sheet|
  #     if (sheet == "LP Open Acces ARMs")
  #      sheet_data = @xlsx.sheet(sheet)
  #       @adjustment_hash = {}
  #       @program_ids = []
  #       @fixed_data = []
  #       @sub_data = []
  #       @unit_data = []
  #       primary_key = ''
  #       secondry_key = ''
  #       misc_adj_key = ''
  #       term_key = ''
  #       ltv_key = ''
  #       cltv_key = ''
  #       misc_key = ''
  #       fixed_key = ''
  #       sub_data = ''
  #       key = ''
  #       @sheet = sheet
  #       (1..35).each do |r|
  #         row = sheet_data.row(r)
  #         if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
  #           rr = r + 1
  #           max_column_section = row.compact.count - 1
  #           (0..max_column_section).each do |max_column|
  #             cc = 3 + max_column*6 # (3 / 9 / 15)
  #             begin
  #               # title
  #               @title = sheet_data.cell(r,cc)

  #               # term
  #               term = nil
  #               program_heading = @title.split
  #               if @title.include?("10yr") || @title.include?("10 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("15yr") || @title.include?("15 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("20yr") || @title.include?("20 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("25yr") || @title.include?("25 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               elsif @title.include?("30yr") || @title.include?("30 Yr")
  #                 term = @title.scan(/\d+/)[0]
  #               end

  #               # interest type
  #               if @title.include?("Fixed")
  #                 loan_type = "Fixed"
  #               elsif @title.include?("ARM")
  #                 loan_type = "ARM"
  #               elsif @title.include?("Floating")
  #                 loan_type = "Floating"
  #               elsif @title.include?("Variable")
  #                 loan_type = "Variable"
  #               else
  #                 loan_type = nil
  #               end

  #               # rate arm
  #               if @title.include?("5-1 ARM") || @title.include?("7-1 ARM") || @title.include?("10-1 ARM") || @title.include?("10-1 ARM")
  #                 arm_basic = @title.scan(/\d+/)[0].to_i
  #               end

  #               # conforming
  #               conforming = false
  #               if @title.include?("Freddie Mac") || @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Possible") || @title.include?("Freddie Mac Home Ready")
  #                 conforming = true
  #               end

  #               # freddie_mac
  #               freddie_mac = false
  #               if @title.include?("Freddie Mac")
  #                 freddie_mac = true
  #               end

  #               # fannie_mae
  #               fannie_mae = false
  #               if @title.include?("Fannie Mae") || @title.include?("Freddie Mac Home Ready")
  #                 fannie_mae = true
  #               end

  #               @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
  #               @program_ids << @program.id
  #                # Loan Limit Type
  #               if @title.include?("Non-Conforming")
  #                 @program.loan_limit_type << "Non-Conforming"
  #               end
  #               if @title.include?("Conforming")
  #                 @program.loan_limit_type << "Conforming"
  #               end
  #               if @title.include?("Jumbo")
  #                 @program.loan_limit_type << "Jumbo"
  #               end
  #               if @title.include?("High Balance")
  #                 @program.loan_limit_type << "High Balance"
  #               end
  #               @program.save
  #               @program.update(term: term,loan_type: loan_type,conforming: conforming,freddie_mac: freddie_mac, fannie_mae: fannie_mae, arm_basic: arm_basic, loan_category: sheet)
  #               
  #               @block_hash = {}
  #               key = ''
  #               # main_key = ''
  #               # if @program.term.present?
  #               #   main_key = "Term/LoanType/InterestRate/LockPeriod"
  #               # else
  #               #   main_key = "InterestRate/LockPeriod"
  #               # end
  #               # @block_hash[main_key] = {}
  #               (0..50).each do |max_row|
  #                 @data = []
  #                 (0..4).each_with_index do |index, c_i|
  #                   rrr = rr + max_row
  #                   ccc = cc + c_i
  #                   value = sheet_data.cell(rrr,ccc)
  #                   if (c_i == 0)
  #                     key = value
  #                     @block_hash[key] = {}
  #                   else
  #                     if @program.lock_period.length <= 3
  #                       @program.lock_period << 15*c_i
  #                       @program.save
  #                     end
  #                     @block_hash[key][15*c_i] = value
  #                   end
  #                   @data << value
  #                 end

  #                 if @data.compact.length == 0
  #                   break # terminate the loop
  #                 end
  #               end
  #               if @block_hash.values.first.keys.first.nil?
  #                 @block_hash.values.first.shift
  #               end
  #               @block_hash.delete(nil)
  #               @program.update(base_rate: @block_hash)
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       # Adjustments
  #       (37..71).each do |r|
  #         row = sheet_data.row(r)
  #         @fixed_data = sheet_data.row(39)
  #         @sub_data = sheet_data.row(47)
  #         @unit_data = sheet_data.row(56)
  #         if row.compact.count >= 1
  #           (0..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if value == "Loan Level Price Adjustments: See Adjustment Caps" || value == "Adjustments Applied after Cap"
  #                   primary_key = value
  #                   @adjustment_hash[primary_key] = {}
  #                 end
  #                 if value == "All LP Open Access ARMs" || value == "Subordinate Financing" || value == "Number Of Units" || value == "Loan Size Adjustments"
  #                   secondry_key = value
  #                   @adjustment_hash[primary_key][secondry_key] = {}
  #                 end

  #                 # All LP Open Access ARMs
  #                 if r >= 40 && r<= 45 && cc == 8# && cc <= 19 && cc != 15
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 40 && r<= 45 && cc > 8 && cc != 15 && cc <= 19
  #                   fixed_key = get_value @fixed_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][fixed_key] = value
  #                 end

  #                 # Subordinate Financing Adjustments
  #                 if r >= 48 && r <= 54 && cc == 5
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 48 && r <= 54 && cc == 6
  #                   cltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key] = {}
  #                 end
  #                 if r >= 48 && r<= 54 && cc >= 9 && cc <= 10
  #                   sub_data = get_value @sub_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][cltv_key][sub_data] = value
  #                 end

  #                 # Number Of Units Adjustments
  #                 if r >= 57 && r <= 58 && cc == 3
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 57 && r <= 58 && cc > 3 && cc <= 7
  #                   sub_data = get_value @unit_data[cc-2]
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = {}
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key][sub_data] = value
  #                 end

  #                 # Adjustments Applied after Cap
  #                 if r >= 61 && r <= 67 && cc == 6
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 61 && r <= 67 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end

  #                 # Other Adjustments
  #                 if r >= 69 && r <= 71 && cc == 3
  #                   ltv_key = get_value value
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = {}
  #                 end
  #                 if r >= 69 && r <= 71 && cc == 10
  #                   @adjustment_hash[primary_key][secondry_key][ltv_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #           (12..19).each do |max_column|
  #             cc = max_column
  #             value = sheet_data.cell(r,cc)
  #             begin
  #               if value.present?
  #                 if  value == "Misc Adjusters" || value == "Adjustment Caps"
  #                   key = value
  #                   @adjustment_hash[key] = {}
  #                 end

  #                 # Misc Adjustments
  #                 if r >= 47 && r <= 57 && cc == 15
  #                   if value.include?("Condo")
  #                     misc_key = "Condo=>75.01=>15.01"
  #                   else
  #                     misc_key = value
  #                   end
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 47 && r <= 57 && cc == 19
  #                   @adjustment_hash[key][misc_key] = value
  #                 end

  #                 # Adjustment Caps
  #                 if r >= 62 && r <= 65 && cc == 16
  #                   misc_key = value
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 17
  #                   term_key = get_value value
  #                   @adjustment_hash[key][misc_key][term_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 18
  #                   ltv_key = get_value value
  #                   @adjustment_hash[key][misc_key][term_key][ltv_key] = {}
  #                 end
  #                 if r >= 62 && r <= 65 && cc == 19
  #                   @adjustment_hash[key][misc_key][term_key][ltv_key] = value
  #                 end
  #                 if r >= 67 && r <= 68 && cc == 12
  #                   misc_key = value
  #                   @adjustment_hash[key][misc_key] = {}
  #                 end
  #                 if r >= 67 && r <= 68 && cc == 16
  #                   @adjustment_hash[key][misc_key] = value
  #                 end
  #               end
  #             rescue Exception => e
  #               error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet, error_detail: e.message)
  #               error_log.save
  #             end
  #           end
  #         end
  #       end
  #       make_adjust(@adjustment_hash, @sheet)
  #       create_program_association_with_adjustment(@sheet)
  #     end
  #   end
  #   redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  # end


  