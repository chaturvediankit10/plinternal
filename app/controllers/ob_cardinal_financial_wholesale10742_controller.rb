class ObCardinalFinancialWholesale10742Controller < ApplicationController
  before_action :get_sheet, only: [:programs, :ak]
  before_action :get_program, only: [:single_program, :program_property]
  def index
    file = File.join(Rails.root,  'OB_Cardinal_Financial_Wholesale10742.xls')
    xlsx = Roo::Spreadsheet.open(file)
    begin
      xlsx.sheets.each do |sheet|
        if (sheet == "AK")
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Cardinal Financial"
          @bank = Bank.find_or_create_by(name: @name)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def ak
    file = File.join(Rails.root,  'OB_Cardinal_Financial_Wholesale10742.xls')
    xlsx = Roo::Spreadsheet.open(file)
    xlsx.sheets.each do |sheet|
      if (sheet == "AK")
        sheet_data = xlsx.sheet(sheet)
        @programs_ids = []
        @ltv_data = []
        @sub_data = []
        @lpmi_data = []
        primary_key = ''
        secondary_key = ''
        main_key = ''
        ltv_key = ''
        cltv_key = ''
        new_key = ''
        term_key = ''
        @adjustment_hash = {}
        @cashout_adjustment = {}
        @product_hash = {}
        @subordinate_hash = {}
        @additional_hash = {}
        @lpmi_hash = {}
        @freddie_adjustment_hash = {}
        @relief_cashout_adjustment = {}
        # Fannie Mae Programs
        (71..298).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end

        # Fannie Mae Adjustments
        (353..429).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(356)
          @sub_data = sheet_data.row(386)
          @lpmi_data = sheet_data.row(410)
          if row.compact.count >= 1
            (2..42).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Fannie Mae Loan Level Price Adjustments"
                  primary_key = "FannieMae"
                elsif value == "Lender Paid Mortgage Insurance"
                  primary_key = "LPMI"
                end
                if value == "All Eligible Mortgages - LLPAs for Terms > 15 Years"
                  secondary_key = "RateType/Term/FICO/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @adjustment_hash[main_key] = {}
                elsif value == "All Eligible Mortgages  Cash-Out Refinance  LLPAs"
                  secondary_key = "Cashout/FICO/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @cashout_adjustment[main_key] = {}
                elsif value == "All Eligible Mortgages Product Feature  LLPAs"
                  secondary_key = "Cashput/Feature/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @product_hash[main_key] = {}
                elsif value == "Mortgages with Subordinate Financing4"
                  secondary_key = "FinancingType/FICO/LTV/CLTV"
                  main_key = primary_key + "/" + secondary_key
                  @subordinate_hash[main_key] = {}
                elsif value == "State Adjustment"
                  secondary_key = "State"
                  main_key = primary_key + "/" + secondary_key
                  @additional_hash[main_key] = {}
                elsif value == "LPMI Adj. >20yr Term"
                  secondary_key = "Term/LTV"
                  term_key = ">20"
                  main_key = primary_key + "/" + secondary_key
                  @lpmi_hash[main_key] = {}
                  @lpmi_hash[main_key][term_key] = {}
                elsif value == "LPMI Adj. ≤ 20yr Term"
                  secondary_key = "Term/LTV"
                  term_key = "≤ 20"
                  main_key = primary_key + "/" + secondary_key
                  @lpmi_hash[main_key] = {}
                  @lpmi_hash[main_key][term_key] = {}
                end

                # All Eligible Mortgages - LLPAs for Terms > 15 Years
                if r >= 357 && r <= 365 && cc == 9
                  ltv_key = get_value value
                  @adjustment_hash[main_key][ltv_key] = {}
                end
                if r >= 357 && r <= 365 && cc >= 18 && cc <= 44
                  ltv_data =  get_value @ltv_data[cc-2]
                  @adjustment_hash[main_key][ltv_key][ltv_data] = {}
                  @adjustment_hash[main_key][ltv_key][ltv_data] = value
                end

                # All Eligible Mortgages  Cash-Out Refinance  LLPAs
                if r >= 367 && r <= 373 && cc == 9
                  ltv_key = get_value value
                  @cashout_adjustment[main_key][ltv_key] = {}
                end
                if r >= 367 && r <= 373 && cc >= 18 && cc <= 44
                  ltv_data =  get_value @ltv_data[cc-2]
                  @cashout_adjustment[main_key][ltv_key][ltv_data] = {}
                  @cashout_adjustment[main_key][ltv_key][ltv_data] = value
                end

                # All Eligible Mortgages Product Feature  LLPAs
                if r >= 375 && r <= 382 && cc == 9
                  if value == "High Balance Purchase or Rate/Term Refi"
                    secondary_key = "HighBalance/LoanPurpose/Feature/LTV"
                    main_key = primary_key + "/" + secondary_key
                    ltv_key = get_value value
                    @product_hash[main_key] = {}  
                    @product_hash[main_key][ltv_key] = {}
                  elsif value == "High Balance Cash-Out Refi"
                    secondary_key = "HighBalance/Cashout/Feature/LTV"
                    main_key = primary_key + "/" + secondary_key
                    ltv_key = get_value value
                    @product_hash[main_key] = {}  
                    @product_hash[main_key][ltv_key] = {}
                  elsif value == "High Balance ARM2"
                    secondary_key = "HighBalance/LoanType"    
                    main_key = primary_key + "/" + secondary_key
                    ltv_key = get_value value
                    @product_hash[main_key] = {}  
                    @product_hash[main_key][ltv_key] = {}
                  else
                    ltv_key = get_value value
                    @product_hash[main_key][ltv_key] = {}
                  end
                end
                if r >= 375 && r <= 382 && cc >= 18 && cc <= 44
                  ltv_data =  get_value @ltv_data[cc-2]
                  @product_hash[main_key][ltv_key][ltv_data] = {}
                  @product_hash[main_key][ltv_key][ltv_data] = value
                end

                # subordinate adjustment
                if r == 387 && cc == 6
                  new_key = value
                  @subordinate_hash[main_key][new_key] = {}
                end
                if r == 387 && cc == 12
                  @subordinate_hash[main_key][new_key] = value
                end
                if r >= 388 && r <= 392 && cc == 6
                  ltv_key = get_value value
                  @subordinate_hash[main_key][ltv_key] = {}
                end
                if r >= 388 && r <= 392 && cc == 9
                  cltv_key = get_value value
                  @subordinate_hash[main_key][ltv_key][cltv_key] = {}
                end
                if r >= 388 && r <= 392 && cc >= 12 && cc <= 15
                  sub_data = get_value @sub_data[cc-2]
                  @subordinate_hash[main_key][ltv_key][cltv_key][sub_data] = {}
                  @subordinate_hash[main_key][ltv_key][cltv_key][sub_data] = value
                end

                # Additional Adjustments5
                if r >= 394 && r <= 398 && cc == 6
                  if value == "R/T or CO Refinance"
                    secondary_key = "LoanType/RefinanceOption/LTV"
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  elsif value == "Escrow Waiver FICO < 700"
                    secondary_key = "EscrowWaiver/FICO"
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  elsif value == "Escrow Waiver CA FICO < 700"
                    secondary_key = "CA/EscrowWaiver/FICO"
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  elsif value == "ARM > 90 LTV"
                    secondary_key = "LoanType/LTV"     
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  elsif value == "90 Day (Add to 60 Day)"
                    secondary_key = "LockPeriod"
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  end
                end
                if r >= 394 && r <= 398 && cc == 14
                  @additional_hash[main_key] = value
                end
                if r == 400 && cc == 2
                  if value == "Max Net Rebate"
                    secondary_key = "Max/Net/Rebate"
                    main_key = primary_key + "/" + secondary_key
                    @additional_hash[main_key] = {}
                  end
                end
                if r == 401 && cc == 2
                  @additional_hash[main_key] = value
                end
                if r == 404 && cc == 2
                  ltv_key = value
                  @additional_hash[main_key][ltv_key] = {}
                end
                if r == 404 && cc == 10
                  @additional_hash[main_key][ltv_key] = value
                end

                # Lender Paid Mortgage Insurance
                if r >= 411 && r <= 416 && cc == 7
                  ltv_key = get_value value
                  # @lpmi_hash[main_key][term_key] = {}
                  @lpmi_hash[main_key][term_key][ltv_key] = {}
                end
                if r >= 411 && r <= 416 && cc == 11
                  cltv_key = get_value value.to_s
                  @lpmi_hash[main_key][term_key][ltv_key][cltv_key] = {}
                end
                if r >= 411 && r <= 416 && cc >= 15 && cc <= 33
                  lpmi_key = get_value @lpmi_data[cc-2]
                  @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = {}
                  @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = value
                end
                # if r >= 418 && r <= 422 && cc == 7
                #   term_key = "≤20" 
                #   ltv_key = get_value value
                #   @lpmi_hash[main_key][term_key] = {}
                #   @lpmi_hash[main_key][term_key][ltv_key] = {}
                # end
                # if r >= 418 && r <= 422 && cc == 11
                #   cltv_key = get_value value.to_s
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key] = {}
                # end
                # if r >= 418 && r <= 422 && cc >= 15 && cc <= 33
                #   lpmi_key = get_value @lpmi_data[cc-2]
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = {}
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = value
                # end
              end
            end
          end
        end
        
        # Freddie programs
        (458..684).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Freddie Adjustments
        (740..835).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(743)
          @sub_data = sheet_data.row(785)
          @lpmi_data = sheet_data.row(816)
          if row.compact.count >= 1
            (2..42).each do |max_column|
              cc = max_column
              value = sheet_data.cell(r,cc)
              if value.present?
                if value == "Freddie Mac Loan Level Price Adjustments"
                  primary_key = "FreddieMac"
                elsif value == "Lender Paid Mortgage Insurance"
                  primary_key = "LPMI"
                end
                if value == "All Eligible Mortgages - Other Than Relief Refinance Mortgages - LLPAs for Terms > 15 Years"
                  secondary_key = "RateType/Term/FICO/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @freddie_adjustment_hash[main_key] = {}
                elsif value == "All Eligible Mortgages - Relief Refinance Mortgages - LLPAs for Terms > 15 Years"
                  secondary_key = "Relief/Cashout/FICO/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @relief_cashout_adjustment[main_key] = {}
                elsif value == "All Eligible Mortgages  Cash-Out Refinance  LLPAs"
                  secondary_key = "Cashout/FICO/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @cashout_adjustment[main_key] = {}
                elsif value == "All Eligible Mortgages Product Feature  LLPAs"
                  secondary_key = "Cashput/Feature/LTV"
                  main_key = primary_key + "/" + secondary_key
                  @product_hash[main_key] = {}
                elsif value == "Mortgages with Subordinate Financing5 - Other Than Relief Refinance Mortgages"
                  secondary_key = "FinancingType/FICO/LTV/CLTV"
                  main_key = primary_key + "/" + secondary_key
                  @subordinate_hash[main_key] = {}
                elsif value == "Mortgages with Subordinate Financing5 - Other Than Relief Refinance Mortgages"
                  secondary_key = "Relief/FinancingType/FICO/LTV/CLTV"
                  main_key = primary_key + "/" + secondary_key
                  @additional_hash[main_key] = {}
                elsif value == "State Adjustment"
                  secondary_key = "State"
                  main_key = primary_key + "/" + secondary_key
                  @additional_hash[main_key] = {}
                elsif value == "LPMI Adj. >20yr Term"
                  secondary_key = "Term/LTV"
                  term_key = ">20"
                  main_key = primary_key + "/" + secondary_key
                  @lpmi_hash[main_key] = {}
                  @lpmi_hash[main_key][term_key] = {}
                elsif value == "LPMI Adj. ≤ 20yr Term"
                  secondary_key = "Term/LTV"
                  term_key = "≤ 20"
                  main_key = primary_key + "/" + secondary_key
                  @lpmi_hash[main_key] = {}
                  @lpmi_hash[main_key][term_key] = {}
                end

                # All Eligible Mortgages - Other Than Relief Refinance Mortgages - LLPAs for Terms > 15 Years
                if r >= 744 && r <= 750 && cc == 10
                  ltv_key = get_value value
                  @freddie_adjustment_hash[main_key][ltv_key] = {}
                end
                if r >= 744 && r <= 750 && cc >= 21 && cc <= 42
                  ltv_data =  get_value @ltv_data[cc-2]
                  @freddie_adjustment_hash[main_key][ltv_key][ltv_data] = {}
                  @freddie_adjustment_hash[main_key][ltv_key][ltv_data] = value
                end

                # All Eligible Mortgages - Relief Refinance Mortgages - LLPAs for Terms > 15 Years
                if r >= 752 && r <= 760 && cc == 10
                  ltv_key = get_value value
                  @relief_cashout_adjustment[main_key][ltv_key] = {}
                end
                if r >= 752 && r <= 760 && cc >= 21 && cc <= 42
                  ltv_data =  get_value @ltv_data[cc-2]
                  @relief_cashout_adjustment[main_key][ltv_key][ltv_data] = {}
                  @relief_cashout_adjustment[main_key][ltv_key][ltv_data] = value
                end

                # All Eligible Mortgages  Cash-Out Refinance  LLPAs
                if r >= 762 && r <= 768 && cc == 10
                  ltv_key = get_value value
                  @cashout_adjustment[main_key][ltv_key] = {}
                end
                if r >= 762 && r <= 768 && cc >= 21 && cc <= 42
                  ltv_data =  get_value @ltv_data[cc-2]
                  @cashout_adjustment[main_key][ltv_key][ltv_data] = {}
                  @cashout_adjustment[main_key][ltv_key][ltv_data] = value
                end

                # # All Eligible Mortgages Product Feature  LLPAs
                # if r >= 375 && r <= 382 && cc == 9
                #   if value == "High Balance Purchase or Rate/Term Refi"
                #     secondary_key = "HighBalance/LoanPurpose/Feature/LTV"
                #     main_key = primary_key + "/" + secondary_key
                #     ltv_key = get_value value
                #     @product_hash[main_key] = {}  
                #     @product_hash[main_key][ltv_key] = {}
                #   elsif value == "High Balance Cash-Out Refi"
                #     secondary_key = "HighBalance/Cashout/Feature/LTV"
                #     main_key = primary_key + "/" + secondary_key
                #     ltv_key = get_value value
                #     @product_hash[main_key] = {}  
                #     @product_hash[main_key][ltv_key] = {}
                #   elsif value == "High Balance ARM2"
                #     secondary_key = "HighBalance/LoanType"    
                #     main_key = primary_key + "/" + secondary_key
                #     ltv_key = get_value value
                #     @product_hash[main_key] = {}  
                #     @product_hash[main_key][ltv_key] = {}
                #   else
                #     ltv_key = get_value value
                #     @product_hash[main_key][ltv_key] = {}
                #   end
                # end
                # if r >= 375 && r <= 382 && cc >= 18 && cc <= 44
                #   ltv_data =  get_value @ltv_data[cc-2]
                #   @product_hash[main_key][ltv_key][ltv_data] = {}
                #   @product_hash[main_key][ltv_key][ltv_data] = value
                # end

                # # subordinate adjustment
                # if r == 387 && cc == 6
                #   new_key = value
                #   @subordinate_hash[main_key][new_key] = {}
                # end
                # if r == 387 && cc == 12
                #   @subordinate_hash[main_key][new_key] = value
                # end
                # if r >= 388 && r <= 392 && cc == 6
                #   ltv_key = get_value value
                #   @subordinate_hash[main_key][ltv_key] = {}
                # end
                # if r >= 388 && r <= 392 && cc == 9
                #   cltv_key = get_value value
                #   @subordinate_hash[main_key][ltv_key][cltv_key] = {}
                # end
                # if r >= 388 && r <= 392 && cc >= 12 && cc <= 15
                #   sub_data = get_value @sub_data[cc-2]
                #   @subordinate_hash[main_key][ltv_key][cltv_key][sub_data] = {}
                #   @subordinate_hash[main_key][ltv_key][cltv_key][sub_data] = value
                # end

                # # Additional Adjustments5
                # if r >= 394 && r <= 398 && cc == 6
                #   if value == "R/T or CO Refinance"
                #     secondary_key = "LoanType/RefinanceOption/LTV"
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   elsif value == "Escrow Waiver FICO < 700"
                #     secondary_key = "EscrowWaiver/FICO"
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   elsif value == "Escrow Waiver CA FICO < 700"
                #     secondary_key = "CA/EscrowWaiver/FICO"
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   elsif value == "ARM > 90 LTV"
                #     secondary_key = "LoanType/LTV"     
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   elsif value == "90 Day (Add to 60 Day)"
                #     secondary_key = "LockPeriod"
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   end
                # end
                # if r >= 394 && r <= 398 && cc == 14
                #   @additional_hash[main_key] = value
                # end
                # if r == 400 && cc == 2
                #   if value == "Max Net Rebate"
                #     secondary_key = "Max/Net/Rebate"
                #     main_key = primary_key + "/" + secondary_key
                #     @additional_hash[main_key] = {}
                #   end
                # end
                # if r == 401 && cc == 2
                #   @additional_hash[main_key] = value
                # end
                # if r == 404 && cc == 2
                #   ltv_key = value
                #   @additional_hash[main_key][ltv_key] = {}
                # end
                # if r == 404 && cc == 10
                #   @additional_hash[main_key][ltv_key] = value
                # end

                # # Lender Paid Mortgage Insurance
                # if r >= 411 && r <= 416 && cc == 7
                #   ltv_key = get_value value
                #   # @lpmi_hash[main_key][term_key] = {}
                #   @lpmi_hash[main_key][term_key][ltv_key] = {}
                # end
                # if r >= 411 && r <= 416 && cc == 11
                #   cltv_key = get_value value.to_s
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key] = {}
                # end
                # if r >= 411 && r <= 416 && cc >= 15 && cc <= 33
                #   lpmi_key = get_value @lpmi_data[cc-2]
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = {}
                #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = value
                # end
                # # if r >= 418 && r <= 422 && cc == 7
                # #   term_key = "≤20" 
                # #   ltv_key = get_value value
                # #   @lpmi_hash[main_key][term_key] = {}
                # #   @lpmi_hash[main_key][term_key][ltv_key] = {}
                # # end
                # # if r >= 418 && r <= 422 && cc == 11
                # #   cltv_key = get_value value.to_s
                # #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key] = {}
                # # end
                # # if r >= 418 && r <= 422 && cc >= 15 && cc <= 33
                # #   lpmi_key = get_value @lpmi_data[cc-2]
                # #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = {}
                # #   @lpmi_hash[main_key][term_key][ltv_key][cltv_key][lpmi_key] = value
                # # end
              end
            end
          end
        end
        # FHA Va Usda programs
        (844..1006).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Non Conforming programs
        (1126..1145).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
        # Jumbo Non Conforming programs
        (1220..1260).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 4)) && !row.include?("Specials") && !row.include?("Max Net Rebate") && !row.include?("State Adjustment") && !row.include?("ALASKA") && !row.include?("> 484.35K") && !row.include?("7. Not applicable if the subordinate financing is an Affordable Second") && !row.include?("> 100k & ≤ 125k ") && !row.include?("8. Applies to Relief Refiance Mortgages Only") && !row.include?("*Approval Required by Credit Committee for all No Credit loans") && !row.include?("Other Specific Adjustments") && !row.include?("*State Adj. Applied After The Cap") && !row.include?("104.5") || (row.include?("Jumbo 5/1 ARM"))
            rr = r + 1 
            max_column_section = row.compact.count - 1
            (0..max_column_section).each_with_index do |max_column, index|
              index = index +1
              cc = 1 + max_column*10 + index# (2 / 13 / 24 / 35)
              @title = sheet_data.cell(r,cc)
              if @title.present?
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                @programs_ids << @program.id
                # Program Property
                program_property @title
                @program.adjustments.destroy_all
              end
              @block_hash = {}
              key = ''
              main_key = ''
              if @program.term.present? 
                main_key = "Term/LoanType/InterestRate/LockPeriod"
              else
                main_key = "InterestRate/LockPeriod"
              end
              @block_hash[main_key] = {}
              (1..50).each do |max_row|
                @data = []
                (0..8).each_with_index do |index, c_i|
                  rrr = rr + max_row +1
                  ccc = cc + c_i
                  value = sheet_data.cell(rrr,ccc)
                  if value.present?
                    if (c_i == 0)
                      key = value
                      @block_hash[main_key][key] = {}
                    else
                      if @program.lock_period.length <= 3
                        @program.lock_period << 15*(c_i/2)
                        @program.save
                      end
                      @block_hash[main_key][key][15*(c_i/2)] = value unless @block_hash[main_key][key].nil?
                    end
                    @data << value
                  end
                end
                if @data.compact.reject { |c| c.blank? }.length == 0
                  break # terminate the loop
                end
              end
              @program.update(base_rate: @block_hash)
            end
          end
        end
      end
    end
    redirect_to programs_ob_cmg_wholesale_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  def get_value value1
    if value1.present?
      if value1.include?("FICO <")
        value1 = "0"+value1.split("FICO").last
      elsif value1.include?("<") || value1.include?("≤")
        value1 = "0"+value1.first(5)
      elsif value1.include?("FICO")
        value1 = value1.split("FICO ").last.first(5)
      elsif value1 == "Investment Property"
        value1 = "Property/Type"
      else
        value1
      end
    end
  end

  private
    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_program
      @program = Program.find(params[:id])
    end

    def program_property value1
      # term
      if @program.program_name.include?("30 Year") || @program.program_name.include?("30Yr") || @program.program_name.include?("30 Yr") || @program.program_name.include?("30/25 Year")
        term = 30
      elsif @program.program_name.include?("20 Year")
        term = 20
      elsif @program.program_name.include?("15 Year")
        term = 15
      elsif @program.program_name.include?("10 Year")
        term = 10
      else
        term = nil
      end

      # Loan-Type
      if @program.program_name.include?("Fixed")
        loan_type = "Fixed"
      elsif @program.program_name.include?("ARM")
        loan_type = "ARM"
      elsif @program.program_name.include?("Floating")
        loan_type = "Floating"
      elsif @program.program_name.include?("Variable")
        loan_type = "Variable"
      else
        loan_type = nil
      end

      # Streamline Vha, Fha, Usda
      fha = false
      va = false
      usda = false
      streamline = false
      full_doc = false
      if @program.program_name.include?("FHA") 
        streamline = true
        fha = true
        full_doc = true
      elsif @program.program_name.include?("VA")
        streamline = true
        va = true
        full_doc = true
      elsif @program.program_name.include?("USDA")
        streamline = true
        usda = true
        full_doc = true
      end
      # Loan Limit Type
      if @program.program_name.include?("Non-Conforming")
        @program.loan_limit_type << "Non-Conforming"
      end
      if @program.program_name.include?("Conforming")
        @program.loan_limit_type << "Conforming"
      end
      if @program.program_name.include?("Jumbo")
        @program.loan_limit_type << "Jumbo"
      end
      if @program.program_name.include?("High Balance")
        @program.loan_limit_type << "High Balance"
      end
      @program.save
      @program.update(term: term, loan_type: loan_type, fha: fha, va: va, usda: usda, full_doc: full_doc, streamline: streamline)
    end
end