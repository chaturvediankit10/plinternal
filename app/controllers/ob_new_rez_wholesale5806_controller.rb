class ObNewRezWholesale5806Controller < ApplicationController
  include ProgramAdj
  method_names = [:government, :programs, :freddie_fixed_rate, :conforming_fixed_rate, :home_possible, :conforming_arms, :lp_open_acces_arms, :lp_open_access_105, :lp_open_access, :du_refi_plus_arms, :du_refi_plus_fixed_rate_105, :du_refi_plus_fixed_rate, :dream_big, :high_balance_extra, :freddie_arms, :jumbo_series_d,:jumbo_series_f, :jumbo_series_h, :jumbo_series_i, :jumbo_series_jqm, :homeready, :homeready_hb]
  before_action :read_sheet, only: method_names + [:index]
  before_action :get_sheet, only: method_names
  # before_action :check_sheet_empty , only: method_names
  before_action :get_program, only: [:single_program, :program_property]

  require 'roo'
  require 'roo-xls'

  def index
    # HardWorker.perform_async(1)
    @banks = Bank.all
    @sheetlist =[]
      @xlsx.sheets.each do |sheet|
        @sheetlist.push(sheet)
        if (sheet == "Cover Zone 1")
          @sheet_name = sheet
          headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @xlsx.sheet(sheet).each_with_index do |row, index|
            current_row = index+1
            if row.include?("Mortgagee Clause (Wholesale)")
              address_index = row.find_index("Mortgagee Clause (Wholesale)")
              @address_a = []
              (1..3).each do |n|
                @address_a << @xlsx.sheet(sheet).row(current_row+n)[address_index]
                if n == 3
                  @zip = @xlsx.sheet(sheet).row(current_row+n)[address_index].split.last
                  @state_code = @xlsx.sheet(sheet).row(current_row+n)[address_index].split[2]
                end
              end
            end
            if (row.include?("Phone") && row.include?("General Contacts"))
              phone_index = row.find_index(headers[0])
              general_contacts_index = row.find_index(headers[1])
              c_row = @xlsx.sheet(sheet).row(current_row+1)
              @name = c_row[general_contacts_index]
              @phone = c_row[phone_index]
            end
          end

          @bank = Bank.find_or_create_by(name: @name)
          @bank.update(phone: @phone, address1: @address_a.join, state_code: @state_code, zip: @zip)
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
  end

  def cover_zone_1
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def heloc
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def smartseries
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def jumbo_series_c
    redirect_to error_page_ob_new_rez_wholesale5806_path
  end

  def government
    @programs_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Government")
        @sheet_name = sheet
        @program_ids = []
        @credit_hash = {}
        @loan_hash = {}
        @hb_hash = {}
        @bpc_loan_hash = {}
        @govt_hash = {}
        @second_hash = {}
        @spe_hash = {}
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        new_key = ''
        new_val = ''
        c_val = ''
        (1..95).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title                                                 
                  @block_hash = {}
                  key = ''
                  (1..23).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
                        end
                        @data << value
                      end
                    end
                    if @data.compact.reject { |c| c.blank? }.length == 0
                      break # terminate the loop
                    end
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (110..136).each do |r|
          row = sheet_data.row(r)
          # @key_data = sheet_data.row(40)
          if (row.compact.count >= 1)
            (0..18).each do |max_column|
              cc = max_column
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @credit_hash["FICO"] = {}
                    @loan_hash["LoanAmount/LoanPurpose"] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"] = {}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]={}
                    @spe_hash["LoanType"] = {}
                    @second_hash["LoanPurpose/LockDay"]={}
                  end
                  if r >= 112 && r <= 120 && cc == 5
                    new_key = get_value value
                    new_val = sheet_data.cell(r,cc+4)
                    @credit_hash["FICO"][new_key] = new_val
                  end
                  if r >= 123 && r <= 127 && cc == 5
                    if value.downcase.include?("conforming")
                      new_key = "300000-Inf"
                    else
                      new_key = get_value value
                    end
                    new_val = sheet_data.cell(r,cc+4)
                    c_val = sheet_data.cell(r,cc+5)
                    @loan_hash["LoanAmount/LoanPurpose"][new_key] = {}
                    @loan_hash["LoanAmount/LoanPurpose"][new_key]["Purchase"] = new_val
                    @loan_hash["LoanAmount/LoanPurpose"][new_key]["Refinance"] = c_val
                  end
                  if r == 128 && cc == 9
                    @hb_hash["LoanSize/LoanPurpose"] = {}
                    @hb_hash["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @hb_hash["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @hb_hash["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = value
                    new_val = sheet_data.cell(r,cc+1)
                    @hb_hash["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @hb_hash["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r >= 129 && r <= 133 && cc == 5
                    new_key = get_value value
                    new_val = sheet_data.cell(r,cc+4)
                    c_val = sheet_data.cell(r,cc+5)
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key] = {}
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key]["Purchase"] = new_val
                    @bpc_loan_hash["VA/LoanAmount/LoanPurpose"]["true"][new_key]["Refinance"] = c_val
                  end
                  if r == 136 && cc == 6
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]["true"]={}
                    @govt_hash["FHA/RefinanceOption/Streamline/VA"]["true"]["IRRRL"]["true"]["true"]=value
                  end
                  if r >= 112 && r <= 125 && cc == 12
                    if value == "30, 45 & 60 Day Lock Purchase Special"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["LoanPurpose/LockDay"]["Purchase"]={}
                      @second_hash["LoanPurpose/LockDay"]["Purchase"][30]=new_val
                      @second_hash["LoanPurpose/LockDay"]["Purchase"][45]=new_val
                      @second_hash["LoanPurpose/LockDay"]["Purchase"][60]=new_val
                    end
                    if value == "FHA Refinances"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/LoanPurpose"]={}
                      @second_hash["FHA/LoanPurpose"]["true"]={}
                      @second_hash["FHA/LoanPurpose"]["true"]["Refinance"]=new_val
                    end
                    if value == "FHA/VA ARM <660"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/LoanType/VA/FICO"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]["true"]={}
                      @second_hash["FHA/LoanType/VA/FICO"]["true"]["ARM"]["true"]["0-660"]=new_val
                    end
                    if value == "90 Day Lock (FRM & Purch Only)"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["LoanType/LoanPurpose/LockDay"]={}
                      @second_hash["LoanType/LoanPurpose/LockDay"]["Fixed"]={}
                      @second_hash["LoanType/LoanPurpose/LockDay"]["Fixed"]["Purchase"]={}
                      @second_hash["LoanType/LoanPurpose/LockDay"]["Fixed"]["Purchase"]["90"]=new_val
                    end
                    if value == "VA Cashout >95 LTV"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["RefinanceOption/VA/LTV"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]["true"]={}
                      @second_hash["RefinanceOption/VA/LTV"]["Cash Out"]["true"]["0-95"]=new_val
                    end
                    if value == "VA - Refinance Credit Score ≥ 620"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["LoanType/VA/FICO"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]["true"]={}
                      @second_hash["LoanType/VA/FICO"]["Refinance"]["true"]["0-620"]=new_val
                    end
                    if value == "VA - All Loan Purposes - Credit Score < 620"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["VA/FICO"]={}
                      @second_hash["VA/FICO"]["true"]={}
                      @second_hash["VA/FICO"]["true"]["0-620"]=new_val
                    end
                    if value == "VA - IRRRL - Investment Property"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["VA/LoanPurpose/RefinanceOption"]={}
                      @second_hash["VA/LoanPurpose/RefinanceOption"]["true"]={}
                      @second_hash["VA/LoanPurpose/RefinanceOption"]["true"]["Refinance"]={}
                      @second_hash["VA/LoanPurpose/RefinanceOption"]["true"]["Refinance"]["IRRRL"]=new_val
                    end
                    if value == "Manufactured Home (FHA Only)"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/PropertyType"]={}
                      @second_hash["FHA/PropertyType"]["true"]={}
                      @second_hash["FHA/PropertyType"]["true"]["Manufactured Home"]=new_val
                    end
                    if value == "High Balance - 15 Yr Term\n(Adjusting 15 Yr Conforming Pricing - FHA/VA ONLY"
                      new_val = sheet_data.cell(r,cc+6)
                      @second_hash["FHA/VA/LoanSize/Term"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]["15"]={}
                      @second_hash["FHA/VA/LoanSize/Term"]["true"]["true"]["High-Balance"]["15"]=new_val
                    end
                    if value == "Margin on all Government ARMs"
                      new_val = sheet_data.cell(r,cc+6)
                      new_val = get_value new_val
                      @second_hash["Margin"]={}
                      @second_hash["Margin"]=new_val
                    end
                  end
                  if r >= 132 && r <= 133 && cc == 17
                    new_val = sheet_data.cell(r,cc+1)
                    @spe_hash["LoanType"]["fixed"] = new_val if value == "Fixed"
                    @spe_hash["LoanType"]["ARM"] = new_val if value == "ARM"
                  end

                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@spe_hash,@hb_hash,@credit_hash,@bpc_loan_hash,@govt_hash,@loan_hash,@second_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end

    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def freddie_fixed_rate
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Freddie Fixed Rate")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        main_key = ''
        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10yr Super Conforming"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title 
                @program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      # first_row[c_i]
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (120..175).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(122)
          @lpmi = sheet_data.row(137)
          @fico = sheet_data.row(158)
          @unit_data = sheet_data.row(155)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 123 && r <= 130 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 123 && r <= 130 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 132 && r <= 135 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 132 && r <= 135 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 138 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 138 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 139 && r <= 142 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 139 && r <= 142 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r == 143 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"] = {}
                  end
                  if r == 143 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = value
                  end
                  if r >= 145 && r <= 148 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 145 && r <= 148 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 150 && r <= 153 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 150 && r <= 153 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 156 && r <= 157 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 156 && r <= 157 && cc >= 9 && cc <= 11
                    unit_data = get_value @unit_data[cc-2]
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = {}
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = value
                  end
                  if r >= 159 && r <= 162 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 159 && r <= 162 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 159 && r <= 162 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 163 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 164 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 165 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 166 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 167 && r <= 169 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 156 && r <= 162 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 156 && r <= 162 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 156 && r <= 162 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 162 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 162 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 162 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 163 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 164 && cc == 17
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"]["Rate and Term"] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Purchase"]["Rate and Term"] = new_val
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["Super Conforming"]["Refinance"]["Rate and Term"] = new_val
                  end
                  if r == 165 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["Super Conforming"]["Cash Out"] = value
                  end
                  if r == 172 && cc == 17
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r == 173 && cc == 17
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    # create adjustment for each program
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def conforming_fixed_rate
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conforming Fixed Rate")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..118).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Fannie Mae 10yr High Balance"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title 
                program_ids << @program.id
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (120..173).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(122)
          @lpmi = sheet_data.row(137)
          @fico = sheet_data.row(155)
          @unit_data = sheet_data.row(155)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 123 && r <= 130 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 123 && r <= 130 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 133 && r <= 135 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 133 && r <= 135 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 138 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 138 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 139 && r <= 142 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 139 && r <= 142 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 145 && r <= 148 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 145 && r <= 148 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 150 && r <= 153 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 150 && r <= 153 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 156 && r <= 161 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 156 && r <= 161 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 156 && r <= 161 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 162 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 163 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r == 164 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 165 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 166 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r >= 167 && r <= 169 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 156 && r <= 162 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 156 && r <= 162 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 156 && r <= 162 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 162 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 162 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 162 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 163 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 169 && cc == 17
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r == 170 && cc == 17
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["90"] = value
                  end
                  if r == 171 && cc == 17
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"]["Rate and Term"] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Purchase"]["Rate and Term"] = new_val
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption"]["High-Balance"]["Refinance"]["Rate and Term"] = new_val
                  end
                  if r == 172 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption"]["High-Balance"]["Cash Out"] = value
                  end               
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end

    # create adjustment for each program
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def home_possible
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Home Possible")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                #title
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title 
                @program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (78..134).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(80)
          @lpmi = sheet_data.row(92)
          @fico = sheet_data.row(114)
          @unit = sheet_data.row(123)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 81 && r <= 88 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 81 && r <= 88 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r == 93 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 93 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 94 && r <= 95 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 94 && r <= 95 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 97 && r <= 100 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 97 && r <= 100 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 102 && r <= 105 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 102 && r <= 105 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r == 107 && cc == 4
                    @property_hash["PropertyType/LTV"]["Manufactured Home"] = {}
                  end
                  if r == 107 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["PropertyType/LTV"]["Manufactured Home"][ltv_key] = {}
                    @property_hash["PropertyType/LTV"]["Manufactured Home"][ltv_key] = value
                  end
                  if r >= 109 && r <= 112 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key] = {}
                  end
                  if r >= 109 && r <= 112 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["30"][primary_key][ltv_key] = value
                  end
                  if r >= 115 && r <= 118 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 115 && r <= 118 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 115 && r <= 118 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 124 && r <= 125 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 124 && r <= 125 && cc >= 9 && cc <= 10
                    unit_data = get_value @unit[cc-2]
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = {}
                    @property_hash["PropertyType/LTV"][primary_key][unit_data] = value
                  end
                  if r == 127 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 116 && r <= 122 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 116 && r <= 122 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 116 && r <= 122 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 122 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 122 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 122 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 123 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 124 && cc == 19
                    @loan_amount["MiscAdjuster"] = {}
                    @loan_amount["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 125 && cc == 19
                    @loan_amount["PropertyType"] = {}
                    @loan_amount["PropertyType"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType"]["Manufactured Home"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def jumbo_series_d
    @adjustment_hash = {}
    @property_hash = {}
    @state = {}
    primary_key = ''
    ltv_key = ''
    secondry_key = ''
    @programs_ids =[]
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_D")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        (1..22).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (2 / 8 / 14)
              begin
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title
                @programs_ids  << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {} if key.present?
                    else
                      begin
                        @block_hash[key][15*c_i] = value if key.present? &&value.present?
                      rescue Exception => e
                      end
                    end
                    @data << value
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #For Adjustments
        (41..71).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(44)
          if row.count >= 1
            (0..17).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "FICO/LTV Adjustments - Loan Amount ≤ $1MM"
                    @adjustment_hash["LoanAmount/FICO/LTV"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.capitalize}"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["PropertyType/LTV"] = {}
                  end
                  if value == "State Adjustments"
                    @state["State"] = {}
                  end
                  # FICO/LTV Adjustments - Loan Amount ≤ $1MM
                  if r >= 45 && r <= 51 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key] = {}
                  end
                  if r >= 45 && r <= 51 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["0-1000000"][primary_key][ltv_key] = value
                  end
                  # State Adjustments
                  if r >= 45 && r <= 61 && cc == 11
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 13
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  if r >= 45 && r <= 61 && cc == 16
                    secondry_key = value
                    @state["State"][secondry_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @state["State"][secondry_key] = new_val
                  end
                  # FICO/LTV Adjustments - Loan Amount > $1MM
                  if r >= 55 && r <= 61 && cc == 3
                    if value.include?(">")
                      primary_key = value.tr('>= ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = value
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.capitalize}"][primary_key] = {}
                  end
                  if r >= 55 && r <= 61 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.capitalize}"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanAmount/FICO/LTV"]["1000000-#{(Float::INFINITY).to_s.capitalize}"][primary_key][ltv_key] = value
                  end
                  # Max Price
                  if r == 64 && cc == 11
                    @adjustment_hash["LoanType/Term"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = {}
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["20"] = new_val
                    @adjustment_hash["LoanType/Term"]["Fixed"]["30"] = new_val
                  end
                  if r == 65 && cc == 11
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @adjustment_hash["LoanType/Term"]["Fixed"]["15"] = new_val
                  end
                  # Feature Adjustments
                  if r >= 65 && r <= 67 && cc == 2
                    if value == "Investment"
                      primary_key = "Investment Property"
                    else
                      primary_key = value
                    end
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 65 && r <= 67 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 68 && cc == 2
                    @property_hash["RefinanceOption/LTV"] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                  end
                  if r == 68 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][ltv_key] = value
                  end
                  if r >= 69 && r <= 70 && cc == 2
                    primary_key = value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 69 && r <= 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][ltv_key] = value
                  end
                  if r == 70 && cc == 2
                    @property_hash["MiscAdjuster/LTV"] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"] = {}
                  end
                  if r == 70 && cc >= 4 && cc <= 9
                    if @ltv_data[cc-1].include?("<")
                      ltv_key = "0-"+@ltv_data[cc-1].tr('<= ','')
                    else
                      ltv_key = @ltv_data[cc-1]
                    end
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = {}
                    @property_hash["MiscAdjuster/LTV"]["Escrow Waiver - except CA"][ltv_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@state]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def jumbo_series_f
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_F")
        @sheet_name = sheet
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
        sheet_data = @xlsx.sheet(sheet)
        (2..36).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 23)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              begin
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title
                if @title.include?("20/25/30 Yr")
                  term = 2030
                elsif @title.include?("10/15 Yr")
                  term = 1015
                end
                @program.update(term: term)
                @block_hash = {}
                key = ''
                (1..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if value.present?
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash, loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        # Adjustments
        (55..94).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(60)
          @cltv_data2 = sheet_data.row(59)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Purchase Transactions"
                    @adjustment_hash["LoanPurpose/FICO/LTV"] = {}
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"] = {}
                    @state["State"] = {}
                  end
                  if value == "R/T Refinance Transactions"
                    @refinance_hash["RefinanceOption/FICO/LTV"] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  end
                  if value == "Loan Amount Adjustments"
                    @loan_amount["LoanAmount/LTV"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["PropertyType/LTV"] = {}
                  end
                  # Loan Amount Adjustments
                  if r >= 60 && r <= 63 && cc == 15
                    if value.include?("≤")
                      ltv_key = "0-"+value.tr('A-Z≤ $ ','')+"000000"
                    else
                      ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                    end
                    @loan_amount["LoanAmount/LTV"][ltv_key] = {}
                  end
                  if r >= 60 && r <= 63 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = {}
                    @loan_amount["LoanAmount/LTV"][ltv_key][secondry_key] = value
                  end
                  # Purchase Transactions Adjustment
                  if r >= 61 && r <= 65 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key] = {}
                  end
                  if r >= 61 && r <= 65 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanPurpose/FICO/LTV"]["Purchase"][primary_key][secondry_key] = value
                  end
                  # Feature Adjustments
                  if r >= 68 && r <= 73 && cc == 15
                    primary_key = value
                    @property_hash["PropertyType/LTV"][primary_key] = {}
                  end
                  if r >= 68 && r <= 73 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].present? && @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["PropertyType/LTV"][primary_key][secondry_key] = {}
                    @property_hash["PropertyType/LTV"][primary_key][secondry_key] = value
                  end
                  # R/T Refinance Transactions Adjustment
                  if r >= 69 && r <= 73 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key] = {}
                  end
                  if r >= 69 && r <= 73 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Rate and Term"][primary_key][secondry_key] = value
                  end
                  # # C/O Refinance Transactions Adjustment
                  if r >= 77 && r <= 81 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 77 && r <= 81 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = {}
                    @refinance_hash["RefinanceOption/FICO/LTV"]["Cash Out"][primary_key][secondry_key] = value
                  end
                  # State Adjustments
                  if r == 86 && cc == 3
                    @state["State"]["FL"] = {}
                    @state["State"]["NV"] = {}
                  end
                  if r ==86 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["State"]["FL"][secondry_key] = {}
                    @state["State"]["NV"][secondry_key] = {}
                    @state["State"]["FL"][secondry_key] = value
                    @state["State"]["NV"][secondry_key] = value
                  end
                  if r == 87 && cc == 3
                    @state["State"]["CA"] = {}
                  end
                  if r ==87 && cc > 3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["State"]["CA"][secondry_key] = {}
                    @state["State"]["CA"][secondry_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def jumbo_series_h
    @program_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_H")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @sheet = sheet
        @adjustment_hash = {}
        @property_hash = {}
        primary_key = ''
        ltv_key = ''
        secondry_key = ''
        (2..86).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Jumbo Series H Product and Pricing"))
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8)/ (r == 28)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 4 + max_column*6 # (4 / 10 / 16/ 22)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
                  @program_ids << @program.id
                  
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        if @program.lock_period.length <= 3
                          @program.lock_period << 15*c_i
                          @program.save
                        end
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end

                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        #For Adjustments
        (68..97).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(71)
          @credit_data = sheet_data.row(84)
          if row.count >= 1
            (0..20).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Jumbo Series H - Adjustments"
                    @adjustment_hash["LoanSize/LoanType/State/Term"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/FICO/LTV"] = {}
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"] = {}
                    @property_hash["RefinanceOption/LTV"] = {}
                    @property_hash["RefinanceOption/LTV"]["Cash Out"] = {}
                  end
                  if r >= 72 && r <= 81 && cc == 12
                    primary_key = value
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 72 && r <= 81 && cc >= 13 && cc <= 17
                    ltv_data = @ltv_data[cc-2].tr('A-Za-z ','')
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/State/Term"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 72 && r <= 81 && cc == 12
                    primary_key = value
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 72 && r <= 81 && cc >= 18 && cc <= 20
                    ltv_data = @ltv_data[cc-2].tr('A-Za-z ','').split('/').first
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/State/ArmBasic"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r >= 85 && r <= 90 && cc == 12
                    primary_key = get_value value
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key] = {}
                  end
                  if r >= 85 && r <= 90 && cc >= 13 && cc <= 17
                    credit_data = get_value @credit_data[cc-2]
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key][credit_data] = {}
                    @property_hash["LoanSize/FICO/LTV"]["Jumbo"][primary_key][credit_data] = value
                  end
                  if r == 93 && cc == 17
                    @property_hash["LoanAmount"] = {}
                    @property_hash["LoanAmount"]["100000-Inf"] = {}
                    @property_hash["LoanAmount"]["100000-Inf"] = value
                  end
                  if r == 94 && cc == 17
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2nd Home"] = {}
                    @property_hash["PropertyType"]["2nd Home"] = value
                  end
                  if r >= 95 && r <= 97 && cc == 15
                    primary_key = get_value value
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][primary_key] = {}
                    cc = cc + 2
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["RefinanceOption/LTV"]["Cash Out"][primary_key] = new_val
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash]
        make_adjust(adjustment,@sheet_name) 
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def jumbo_series_i
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_I")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @property_hash = {}
        @state = {}
        primary_key = ''
        ltv_key = ''
        secondry_key = ''
        main_key = ''
        @sheet = sheet
        # programs
        (2..32).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
                  @program_ids << @program.id
                  
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end
                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        #For Adjustments
        (34..73).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(40)
          if row.count >= 1
            (0..19).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Jumbo Series I Adjustments"
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"] = {}
                    @state["LoanSize/LoanType/Term"] = {}
                    @state["LoanSize/LoanType/Term"]["Jumbo"] = {}
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"] = {}
                    @state["LoanSize/LoanType/ArmBasic"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"] = {}
                  end
                  if r >= 41 && r <= 45 && cc == 3
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 41 && r <= 45 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 41 && r <= 45 && cc == 12
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 41 && r <= 45 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r >= 50 && r <= 53 && cc == 3
                    primary_key = convert_range value
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 50 && r <= 53 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["Fixed"][primary_key][ltv_data] = value
                  end
                  if r >= 50 && r <= 53 && cc == 12
                    primary_key = convert_range value
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 50 && r <= 53 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/LoanAmount/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r == 58 && cc == 3
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"] = {}
                  end
                  if r == 58 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2nd Home"][ltv_data] = value
                  end
                  if r == 58 && cc == 12
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"] = {}
                  end
                  if r == 58 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2nd Home"][ltv_data] = value
                  end
                  if r == 59 && cc == 3
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"] = {}
                  end
                  if r == 59 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = value
                  end
                  if r == 59 && cc == 12
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"] = {}
                  end
                  if r == 59 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = value
                  end
                  if r == 60 && cc == 3
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                  end
                  if r == 60 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["Fixed"]["Cash Out"][ltv_data] = value
                  end
                  if r == 60 && cc == 12
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"] = {}
                  end
                  if r == 60 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = value
                  end
                  if r == 61 && cc == 3
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"] = {}
                  end
                  if r == 61 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["Fixed"]["2-4 Unit"][ltv_data] = value
                  end
                  if r == 61 && cc == 12
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"] = {}
                  end
                  if r == 61 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/PropertyType/LTV"]["Jumbo"]["ARM"]["2-4 Unit"][ltv_data] = value
                  end
                  if r == 62 && cc == 3
                    @property_hash["LoanSize/LoanType/DTI/LTV"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"] = {}
                  end
                  if r == 62 && cc >= 5 && cc <= 10
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["Fixed"]["40%"][ltv_data] = value
                  end
                  if r == 62 && cc == 12
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"] = {}
                  end
                  if r == 62 && cc >= 14 && cc <= 19
                    ltv_data = get_value @ltv_data[cc-2]
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"][ltv_data] = {}
                    @property_hash["LoanSize/LoanType/DTI/LTV"]["Jumbo"]["ARM"]["40%"][ltv_data] = value
                  end
                  if r == 66 && cc == 7
                    @state["MiscAdjuster/State"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["CA"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NC"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["DC"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NY"] = {}
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["CA"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NC"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["DC"] = value
                    @state["MiscAdjuster/State"]["Escrow Waiver"]["NY"] = value
                  end
                  if r == 66 && cc ==17
                    @state["LoanSize/LoanPurpose/LoanType/Term"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"]["30"] = {}
                    @state["LoanSize/LoanPurpose/LoanType/Term"]["Jumbo"]["Purchase"]["Fixed"]["30"] = value
                  end
                  if r >= 72 && r <= 73 && cc == 3
                    primary_key = value.tr('A-Za-z ','')
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @state["LoanSize/LoanType/Term"]["Jumbo"]["Fixed"][primary_key] = new_val
                  end
                  if r == 72 && cc == 17
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["7"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["10"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["7"] = value
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["10"] = value
                  end
                  if r == 73 && cc == 17
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["5"] = {}
                    @state["LoanSize/LoanType/ArmBasic"]["Jumbo"]["ARM"]["5"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@state]
        make_adjust(adjustment,@sheet_name)        
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def jumbo_series_jqm
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Jumbo Series_JQM")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @refinance_hash = {}
        @loan_amount = {}
        @state = {}
        @property_hash = {}
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        @sheet = sheet
        (2..60).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8)/ (r == 21)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 6 + max_column*6 # (6 / 12 / 18)
              begin
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.update_fields p_name
                  program_property @title
                  if @title.include?("20/25/30 Yr")
                    term = 2030
                  elsif @title.include?("10/15 Yr")
                    term = 1015
                  end           
                  @program_ids << @program.id
                  @program.update(term: term)
                  
                  @block_hash = {}
                  key = ''
                  (0..50).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if (c_i == 0)
                        key = value
                        @block_hash[key] = {}
                      else
                        @block_hash[key][15*c_i] = value
                      end
                      @data << value
                    end
                    if @data.compact.length == 0
                      break #terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil?
                    @block_hash.values.first.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (64..99).each do |r|
          row = sheet_data.row(r)
          @cltv_data = sheet_data.row(67)
          @cltv_data2 = sheet_data.row(66)
          @max_price_data = sheet_data.row(94)
          if row.compact.count >= 1
            (3..25).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "Purchase Transactions"
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"] = {}
                    @state["LoanSize/State/LTV"] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"] = {}
                  end
                  if value == "R/T Refinance Transactions"
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"] = {}
                  end
                  if value == "Loan Amount Adjustments"
                    @loan_amount["LoanSize/LoanAmount/LTV"] = {}
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"] = {}
                  end
                  if value == "Feature Adjustments"
                    @property_hash["LoanSize/PropertyType/LTV"] = {}
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"] = {}
                  end
                  # Loan Amount Adjustments
                  if r >= 67 && r <= 70 && cc == 15
                    if value.include?("≤")
                      ltv_key = "0-"+value.tr('A-Z≤ $ ','')+"000000"
                    else
                      ltv_key = (value.tr('A-Z$ ','').split("-").first.to_f*1000000).to_s+"-"+(value.tr('A-Z$ ','').split("-").last.to_f*1000000).to_s
                    end
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key] = {}
                  end
                  if r >= 67 && r <= 70 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key][secondry_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LTV"]["Jumbo"][ltv_key][secondry_key] = value
                  end
                  # Purchase Transactions Adjustment
                  if r >= 68 && r <= 74 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key] = {}
                  end
                  if r >= 68 && r <= 74 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanPurpose/FICO/LTV"]["Jumbo"]["Purchase"][primary_key][secondry_key] = value
                  end
                  # Feature Adjustments
                  if r >= 75 && r <= 78 && cc == 15
                    if value == "Investment"
                      primary_key = "Investment Property"
                    else
                      primary_key = value
                    end
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key] = {}
                  end
                  if r >= 75 && r <= 78 && cc > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key][secondry_key] = {}
                    @property_hash["LoanSize/PropertyType/LTV"]["Jumbo"][primary_key][secondry_key] = value
                  end
                  if r == 79 && cc == 15
                    @property_hash["LoanSize/MiscAdjuster/LTV"] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"] = {}
                  end
                  if r == 79 && cc  > 15 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"][secondry_key] = {}
                    @property_hash["LoanSize/MiscAdjuster/LTV"]["Jumbo"]["Escrow Waiver"][secondry_key] = value
                  end
                  # R/T Refinance Transactions Adjustment
                  if r >= 78 && r <= 84 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key] = {}
                  end
                  if r >= 78 && r <= 84 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key][secondry_key] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Rate and Term"][primary_key][secondry_key] = value
                  end
                  # # C/O Refinance Transactions Adjustment
                  if r >= 88 && r <= 94 && cc == 3
                    if value.include?("≥")
                      primary_key = value.tr('≥ ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      primary_key = get_value value
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 88 && r <= 94 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key][secondry_key] = {}
                    @refinance_hash["LoanSize/RefinanceOption/FICO/LTV"]["Jumbo"]["Cash Out"][primary_key][secondry_key] = value
                  end
                  # State Adjustments
                  if r == 99 && cc == 3
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"] = {}
                  end
                  if r ==99 && cc >3 && cc <= 13
                    if @cltv_data[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data[cc-2]
                    end
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"][secondry_key] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"][secondry_key] = {}
                    @state["LoanSize/State/LTV"]["Jumbo"]["FL"][secondry_key] = value
                    @state["LoanSize/State/LTV"]["Jumbo"]["NV"][secondry_key] = value
                  end
                  if r >= 85 && r <= 87 && cc == 15
                    primary_key = value.tr('A-Za-z ','')
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key] = {}
                  end
                  if r >= 85 && r <= 87 && cc >= 16 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Hybrid"][primary_key][secondry_key] = value
                  end
                  if r >= 88 && r <= 90 && cc == 15        
                    if value == "20 yr Fixed\n(add 30 yr)"
                      primary_key = "20-30"
                    else
                      primary_key = value.tr('A-Za-z ','')
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key] = {}
                  end
                  if r >= 88 && r <= 90 && cc >= 16 && cc <= 25
                    if @cltv_data2[cc-2].include?("≤")
                      secondry_key = "0-"+@cltv_data2[cc-2].tr('≤ ','')
                    else
                      secondry_key = get_value @cltv_data2[cc-2]
                    end
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key][secondry_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV"]["Jumbo"]["Fixed"][primary_key][secondry_key] = value
                  end
                  if r == 96 && cc == 16
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"] = {}

                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"] = {}
                  end
                  if r == 96 && cc >= 19 && cc <= 20
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','')
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["0-100000"]["Fixed"][max_key] = value
                  end
                  if r == 96 && cc >= 21 && cc <= 23
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','').split("/").first
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["0-100000"]["ARM"][max_key] = value
                  end
                  if r == 97 && cc == 16
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"] = {}

                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"] = {}
                  end
                  if r == 97 && cc >= 19 && cc <= 20
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','')
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/Term"]["Jumbo"]["100000-Inf"]["Fixed"][max_key] = value
                  end
                  if r == 97 && cc >= 21 && cc <= 23
                    max_key = @max_price_data[cc-2].tr('A-Za-z ','').split("/").first
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"][max_key] = {}
                    @loan_amount["LoanSize/LoanAmount/LoanType/ArmBasic"]["Jumbo"]["100000-Inf"]["ARM"][max_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@refinance_hash,@loan_amount,@state,@property_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def dream_big
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "Dream Big")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @jumbo_adjustment = {}
        @cash_out = {}
        @other_hash = {}
        @program_ids = []
        primary_key = ''
        fixed_key = ''
        ltv_key = ''
        program_heading = sheet_data.cell(1,10)
        (1..33).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("Dream Big Jumbo"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet + " " + program_heading
                @program.update_fields p_name
                program_property @title
                if @title.include?("20/25/30")
                  term = 2030
                  @program.update(term: term)
                end
                @program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (38..62).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(39)
          @ltv_arm_data = sheet_data.row(54)
          if row.compact.count >= 1
            (0..18).each do |cc|
              value = sheet_data.cell(r,cc)
              begin
                if value.present?
                  if value == "LTV Based Adjustments for 20/25/30 Yr Fixed Jumbo Products"
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"] = {}
                  end
                  if value == "LTV Based Adjustments for 15 Yr Fixed and All ARM Jumbo Products"
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"] = {}
                  end
                  if r >= 40 && r <= 45 && cc == 3
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key] = {}
                  end
                  if r >= 40 && r <= 45 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key][ltv_data] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["20"][primary_key][ltv_data] = value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["25"][primary_key][ltv_data] = value
                    @adjustment_hash["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["30"][primary_key][ltv_data] = value
                  end
                  if r == 46 && cc == 2
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"] = {}
                  end
                  if r == 46 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["30"][ltv_data] = value
                  end
                  if r == 47 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"] = {}
                  end
                  if r == 47 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["30"][ltv_data] = value
                  end
                  if r == 48 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"] = {}
                  end
                  if r == 48 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Rate and Term"]["30"][ltv_data] = value
                  end
                  if r == 50 && cc == 2
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["20"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["25"] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["30"] = {}
                  end
                  if r == 50 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_data[cc-2]
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["20"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["25"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["30"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["20"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["25"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/PropertyType/Term/LTV"]["Jumbo"]["Fixed"]["Non-Owner Occupied"]["30"][ltv_data] = value
                  end
                  if r >= 55 && r <= 60 && cc == 3
                    primary_key = get_value value
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key] = {}
                  end
                  if r >= 55 && r <= 60 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key][ltv_data] = {}
                    @jumbo_adjustment["LoanSize/LoanType/Term/FICO/LTV"]["Jumbo"]["Fixed"]["15"][primary_key][ltv_data] = value
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = {}
                    @jumbo_adjustment["LoanSize/LoanType/FICO/LTV"]["Jumbo"]["ARM"][primary_key][ltv_data] = value
                  end
                  if r == 61 && cc == 2
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"] = {}
                  end
                  if r == 61 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/Term/LTV"]["Jumbo"]["Fixed"]["Purchase"]["15"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/LoanPurpose/LTV"]["Jumbo"]["ARM"]["Purchase"][ltv_data] = value
                  end
                  if r == 62 && cc == 2
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"] = {}
                  end
                  if r == 62 && cc >= 4 && cc <= 14
                    ltv_data = get_value @ltv_arm_data[cc-2]
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/Term/LTV"]["Jumbo"]["Fixed"]["Cash Out"]["15"][ltv_data] = value
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = {}
                    @cash_out["LoanSize/LoanType/RefinanceOption/LTV"]["Jumbo"]["ARM"]["Cash Out"][ltv_data] = value
                  end
                  if r == 44 && cc == 18
                    @other_hash["LoanType/Term"] = {}
                    @other_hash["LoanType/Term"]["Fixed"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["20"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["25"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["30"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["20"] = value
                    @other_hash["LoanType/Term"]["Fixed"]["25"] = value
                    @other_hash["LoanType/Term"]["Fixed"]["30"] = value
                  end
                  if r == 45 && cc == 18
                    @other_hash["LoanType/Term"]["Fixed"]["15"] = {}
                    @other_hash["LoanType/Term"]["Fixed"]["15"] = value
                  end
                  if r == 46 && cc == 18
                    @other_hash["LoanType/ArmBasic"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["5"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["5"] = value
                  end
                  if r == 47 && cc == 18
                    @other_hash["LoanType/ArmBasic"]["ARM"]["7"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["7"] = value
                  end
                  if r == 48 && cc == 18
                    @other_hash["LoanType/ArmBasic"]["ARM"]["10"] = {}
                    @other_hash["LoanType/ArmBasic"]["ARM"]["10"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@jumbo_adjustment,@cash_out,@other_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def high_balance_extra
    @programs_ids = []
    @xlsx.sheets.each do |sheet|
      if (sheet == "High Balance Extra")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @program_ids = []
        @adjustment_hash = {}
        @sub_hash = {}
        @cash_out = {}
        @bal_data = []
        @sub_data = []
        primary_key = ''
        secondry_key = ''
        ltv_key = ''
        cltv_key = ''
        key = ''
        bal_data = ''
        sub_data = ''
        (5..23).each do |r|
          row = sheet_data.row(r)
          if (row.compact.include?("High Balance Extra 30 Yr Fixed"))
            rr = r + 1
            max_column_section = row.compact.count
            (0..max_column_section).each do |max_column|
              cc = 2 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                p_name = @title + " " + sheet
                if @title.present?
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.update_fields p_name
                  program_property @title                         
                  @program_ids << @program.id
                  
                  @block_hash = {}
                  key = ''
                  (0..19).each do |max_row|
                    @data = []
                    (0..4).each_with_index do |index, c_i|
                      rrr = rr + max_row
                      ccc = cc + c_i
                      value = sheet_data.cell(rrr,ccc)
                      if value.present?
                        if (c_i == 0)
                          key = value
                          @block_hash[key] = {}
                        else
                          @block_hash[key][15*c_i] = value
                        end
                        @data << value
                      end
                    end

                    if @data.compact.length == 0
                      break # terminate the loop
                    end
                  end
                  if @block_hash.values.first.keys.first.nil? || @block_hash.keys.first == "Rate"
                    @block_hash.shift
                  end
                  @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (25..44).each do |r|
          row = sheet_data.row(r)
          @bal_data = sheet_data.row(27)
          @sub_data = sheet_data.row(41)
          if row.compact.count >= 1
            (0..9).each do |max_column|
              cc = max_column
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Pricing Adjustments"
                    @adjustment_hash["LoanSize/FICO/LTV"] = {}
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"] = {}
                  end
                  if value == "Cashout (adjustments are cumulative)"
                    @cash_out["RefinanceOption/FICO/LTV"] = {}
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  end
                  if value == "Sub Financing (adjustments are cumulative)"
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  # All High Balance Extra Loans
                  if r >= 28 && r <= 32 && cc == 2
                    if value.include?(">")
                      ltv_key = value.tr('>=','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      ltv_key = get_value value
                    end
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key] = {}
                  end
                  if r >= 28 && r <= 32 && cc > 3 && cc <= 9
                    if @bal_data[cc-2].include?("<")
                      bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                    else
                      bal_data = get_value @bal_data[cc-2]
                    end
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key][bal_data] = {}
                    @adjustment_hash["LoanSize/FICO/LTV"]["High-Balance Extra"][ltv_key][bal_data] = value
                  end
                  # Cashout Adjustments
                  if r >= 34 && r <= 38 && cc == 2
                    if value.include?(">")
                      ltv_key = value.tr('>=','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      ltv_key = get_value value
                    end
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key] = {}
                  end
                  if r >= 34 && r <= 38 && cc > 3 && cc <= 9
                    if @bal_data[cc-2].include?("<")
                      bal_data = "0-"+ @bal_data[cc-2].tr('<= ','')
                    else
                      bal_data = get_value @bal_data[cc-2]
                    end
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = {}
                    @cash_out["RefinanceOption/FICO/LTV"]["Cash Out"][ltv_key][bal_data] = value
                  end

                  # Subordinate Financing Adjustments
                  if r >= 42 && r <= 44 && cc == 2
                    if value.include?("<")
                      ltv_key = "0-"+ value.tr('<= ','')
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key] = {}
                  end
                  if r >= 42 && r <= 44 && cc == 3
                    cltv_key = get_value value
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key] = {}
                  end
                  if r >= 42 && r <= 44 && cc > 3 && cc <= 5
                    if @sub_data[cc-2].include?(">")
                      sub_data = @sub_data[cc-2].tr('>= ','')+"-#{(Float::INFINITY).to_s.capitalize}"
                    else
                      sub_data = get_value @sub_data[cc-2]
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][ltv_key][cltv_key][sub_data] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@sub_hash,@cash_out]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def freddie_arms
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Freddie ARMs")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || (row.include?("Freddie Mac 10-1 ARM (5-2-5) Super Conforming"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                # title
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title    
                @program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (49..99).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(52)
          @lpmi = sheet_data.row(68)
          @fico = sheet_data.row(84)
          @unit = sheet_data.row(81)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/LTV/FICO"] = {}
                    @property_hash["LPMI/LTV/FICO"][true] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"] = {}
                    @property_hash["PropertyType"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"] = {}
                  end
                  if r >= 53 && r <= 59 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key] = {}
                  end
                  if r >= 53 && r <= 59 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = value
                  end
                  if r >= 63 && r <= 66 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 63 && r <= 66 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 69 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 69 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 70 && r <= 73 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 70 && r <= 73 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r == 74 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"] = {}
                  end
                  if r == 74 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Cash Out"][ltv_key] = value
                  end
                  if r >= 76 && r <= 79 && cc == 5
                    primary_key = get_value value
                    @property_hash["LPMI/LTV/FICO"][true][primary_key] = {}
                  end
                  if r >= 76 && r <= 79 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 82 && r <= 83 && cc == 6
                    primary_key = value.split('s').first
                    @property_hash["PropertyType"][primary_key] = {}
                  end
                  if r >= 82 && r <= 83 && cc >= 9 && cc <= 11
                    unit_data = get_value @unit[cc-2]
                    @property_hash["PropertyType"][primary_key][unit_data] = {}
                    @property_hash["PropertyType"][primary_key][unit_data] = value
                  end
                  if r >= 85 && r <= 88 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 85 && r <= 88 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 85 && r <= 88 && cc >= 10 && cc <= 11
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 89 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-inf"] = value
                  end
                  if r == 90 && cc == 11
                    @property_hash["LTV"] = {}
                    @property_hash["LTV"]["90-Inf"] = {}
                    @property_hash["LTV"]["90-Inf"] = value
                  end
                  if r == 91 && cc == 11
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 92 && cc == 11
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 93 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 94 && r <= 96 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 82 && r <= 88 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 82 && r <= 88 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 82 && r <= 88 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 88 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 88 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 88 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 89 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 91 && cc == 19
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["0-75"] = {}
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["0-75"] = value
                  end
                  if r == 92 && cc == 19
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["75-Inf"] = {}
                   @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["Super Conforming"]["Purchase"]["Rate and Term"]["75-Inf"] = value
                  end
                  if r == 93 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["0-75"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["0-75"] = value
                  end
                  if r == 93 && cc == 19
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["75-Inf"] = {}
                    @loan_amount["LoanSize/RefinanceOption/LTV"]["Super Conforming"]["Cash Out"]["75-Inf"] = value
                  end
                  if r == 96 && cc == 19
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def conforming_arms
    @program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "Conforming ARMs")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..47).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet")) || row.compact.include?("Fannie Mae 10-1 ARM (5-2-5) High Balance")
            # r == 7 / 35 / 55
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15)
              begin
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title    
                @program_ids << @program.id  
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end
                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
                # @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (49..99).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(52)
          @lpmi = sheet_data.row(67)
          @fico = sheet_data.row(80)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/LTV/FICO"] = {}
                    @property_hash["LPMI/LTV/FICO"][true] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"] = {}
                    @property_hash["PropertyType"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"] = {}
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"] = {}
                  end
                  if r >= 53 && r <= 60 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key] = {}
                  end
                  if r >= 53 && r <= 60 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/LTV/FICO"]["Conforming"]["ARM"][primary_key][ltv_key] = value
                  end
                  if r >= 63 && r <= 65 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 63 && r <= 65 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 68 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 68 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 69 && r <= 72 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 69 && r <= 72 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 75 && r <= 78 && cc == 5
                    primary_key = get_value value
                    @property_hash["LPMI/LTV/FICO"][true][primary_key] = {}
                  end
                  if r >= 75 && r <= 78 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/LTV/FICO"][true][primary_key][ltv_key] = value
                  end
                  
                  if r >= 81 && r <= 84 && cc == 5
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 81 && r <= 84 && cc == 6
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 81 && r <= 84 && cc >= 9 && cc <= 10
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r >= 87 && r <= 89 && cc == 6
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = {}
                    cc = cc + 3
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = new_val
                  end
                  if r >= 91 && r <= 92 && cc == 5
                    primary_key = value.split('s').first
                    @property_hash["PropertyType"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType"][primary_key] = new_val
                  end
                  if r == 93 && cc == 5
                    @property_hash["PropertyType/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/LTV"]["Condo"]["75-Inf"] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Condo"]["75-Inf"] = new_val
                  end
                  if r == 94 && cc == 9
                    @property_hash["MiscAdjuster"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Full or Taxes Only)"] = value
                  end
                  if r == 95 && cc == 9
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = {}
                    @property_hash["MiscAdjuster"]["CA Escrow Waiver (Insurance Only)"] = value
                  end
                  if r == 96 && cc == 9
                    @property_hash["LTV"] = {}
                    @property_hash["LTV"]["90-Inf"] = {}
                    @property_hash["LTV"]["90-Inf"] = value
                  end
                  if r >= 81 && r <= 87 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 81 && r <= 87 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 81 && r <= 87 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 87 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 87 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 87 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 88 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r >= 89 && r <= 90 && cc == 18
                    primary_key = get_value value
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"][primary_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Rate and Term"][primary_key] = new_val
                  end
                  if r >= 91 && r <= 92 && cc == 18
                    primary_key = get_value value
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"][primary_key] = {}
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose/RefinanceOption/LTV"]["High-Balance"]["Refinance"]["Cash Out"][primary_key] = new_val
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def homeready
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..76).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 # (3 / 9 / 15) 3/8/13
              begin
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title              
                program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: row, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (73..131).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(80)
          @lpmi = sheet_data.row(97)
          @fico = sheet_data.row(112)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @property_hash["PropertyType/LTV"] = {}
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 81 && r <= 88 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 81 && r <= 88 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 91 && r <= 93 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 91 && r <= 93 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 98 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 98 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 99 && r <= 100 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 99 && r <= 100 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 102 && r <= 105 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 102 && r <= 105 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 107 && r <= 110 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 107 && r <= 110 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 113 && r <= 118 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 113 && r <= 118 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 113 && r <= 118 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 119 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r >= 122 && r <= 124 && cc == 7
                    primary_key = get_value value
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = {}
                    cc = cc + 4
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["PropertyType/LTV"]["Non-Owner Occupied"][primary_key] = new_val
                  end
                  if r == 126 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 114 && r <= 120 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 114 && r <= 120 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 114 && r <= 120 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 120 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 120 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 120 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 121 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 122 && cc == 19
                    @loan_amount["MiscAdjuster/LoanPurpose"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = value
                  end
                  if r == 123 && cc == 19
                    @loan_amount["PropertyType/LoanPurpose"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def homeready_hb
    program_ids = []
    @allAdjustments = {}
    @xlsx.sheets.each do |sheet|
      if (sheet == "HomeReady HB")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @property_hash = {}
        @sub_hash = {}
        @loan_amount = {}
        primary_key = ''
        ltv_key = ''
        loan_key = ''
        (1..75).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count > 1) && (row.compact.count <= 3)) && (!row.compact.include?("California Wholesale Rate Sheet"))
            rr = r + 1 # (r == 8) / (r == 36) / (r == 56)
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 3 + max_column*6 
              begin
                @title = sheet_data.cell(r,cc)
                @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                p_name = @title + " " + sheet
                @program.update_fields p_name
                program_property @title   
                program_ids << @program.id
                
                @block_hash = {}
                key = ''
                (0..50).each do |max_row|
                  @data = []
                  (0..4).each_with_index do |index, c_i|
                    rrr = rr + max_row
                    ccc = cc + c_i
                    value = sheet_data.cell(rrr,ccc)
                    if (c_i == 0)
                      key = value
                      @block_hash[key] = {}
                    else
                      @block_hash[key][15*c_i] = value
                    end
                    @data << value
                  end

                  if @data.compact.length == 0
                    break # terminate the loop
                  end
                end
                if @block_hash.values.first.keys.first.nil?
                  @block_hash.values.first.shift
                end
                @block_hash.delete(nil)
                @program.update(base_rate: @block_hash,loan_category: @sheet_name)
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: @sheet_name, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        # Adjustments
        (77..130).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(79)
          @lpmi = sheet_data.row(96)
          @fico = sheet_data.row(111)
          if row.compact.count >= 1
            (0..19).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "Loan Level Price Adjustments"
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"] = {}
                  end
                  if value == "LPMI Adjustments Applied after Cap"
                    @property_hash["LPMI/PropertyType/FICO"] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @property_hash["LoanSize/LoanType/LTV"] = {}
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  end
                  if value == "Adjustments Applied after Cap"
                    @loan_amount["LoanAmount/LoanPurpose"] = {}
                  end
                  if r >= 80 && r <= 87 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key] = {}
                  end
                  if r >= 80 && r <= 87 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = {}
                    @adjustment_hash["LoanSize/LoanType/Term/LTV/FICO"]["Conforming"]["Fixed"]["0-15"][primary_key][ltv_key] = value
                  end
                  if r >= 90 && r <= 92 && cc == 7
                    primary_key = get_value value
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key] = {}
                  end
                  if r >= 90 && r <= 92 && cc >= 10 && cc <= 19
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = {}
                    @adjustment_hash["RefinanceOption/LTV/FICO"]["Cash Out"][primary_key][ltv_key] = value
                  end
                  if r == 97 && cc == 5
                    @property_hash["LPMI/RefinanceOption/FICO"] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"] = {}
                  end
                  if r == 97 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = {}
                    @property_hash["LPMI/RefinanceOption/FICO"][true]["Rate and Term"][ltv_key] = value
                  end
                  if r >= 98 && r <= 99 && cc == 5
                    primary_key = value
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key] = {}
                  end
                  if r >= 98 && r <= 99 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = {}
                    @property_hash["LPMI/PropertyType/FICO"][true][primary_key][ltv_key] = value
                  end
                  if r >= 101 && r <= 104 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key] = {}
                  end
                  if r >= 101 && r <= 104 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["0-20"][primary_key][ltv_key] = value
                  end
                  if r >= 106 && r <= 109 && cc == 6
                    primary_key = get_value value
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key] = {}
                  end
                  if r >= 106 && r <= 109 && cc >= 7 && cc <= 19
                    ltv_key = get_value @lpmi[cc-2]
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = {}
                    @property_hash["LPMI/Term/LTV/FICO"][true]["20-Inf"][primary_key][ltv_key] = value
                  end
                  if r >= 112 && r <= 117 && cc == 6
                    if value.downcase.include?('all')
                      primary_key = "0-Inf"
                    else
                      primary_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key] = {}
                  end
                  if r >= 112 && r <= 117 && cc == 7
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key] = {}
                  end
                  if r >= 112 && r <= 117 && cc >= 10 && cc <= 12
                    cltv_key = get_value @fico[cc-2]
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = {}
                    @sub_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][primary_key][ltv_key][cltv_key] = value
                  end
                  if r == 118 && cc == 11
                    @property_hash["PropertyType"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = {}
                    @property_hash["PropertyType"]["2-4 Unit"] = value
                  end
                  if r == 119 && cc == 11
                    @property_hash["PropertyType/Term/LTV"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = {}
                    @property_hash["PropertyType/Term/LTV"]["Condo"]["15-Inf"]["75-Inf"] = value
                  end
                  if r == 120 && cc == 11
                    @property_hash["PropertyType"]["Manufactured Home"] = {}
                    @property_hash["PropertyType"]["Manufactured Home"] = value
                  end
                  if r == 121 && cc == 11
                    @property_hash["LoanSize/LoanType"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"]["Fixed"] = {}
                    @property_hash["LoanSize/LoanType"]["High-Balance"]["Fixed"] = value
                  end
                  if r >= 122 && r <= 123 && cc == 6
                    primary_key = get_value value
                    primary_key =  primary_key.gsub('--','-')
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][primary_key] = {}
                    cc = cc + 5
                    new_val = sheet_data.cell(r,cc)
                    @property_hash["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][primary_key] = new_val
                  end
                  if r == 125 && cc == 8
                    @property_hash["LoanPurpose/LockDay"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = {}
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["30"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["45"] = value
                    @property_hash["LoanPurpose/LockDay"]["Purchase"]["60"] = value
                  end
                  if r >= 113 && r <= 119 && cc == 15
                    if value.downcase.include?("conforming")
                      loan_key = "300000-Inf"
                    else
                      loan_key = get_value value
                    end
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key] = {}
                  end
                  if r >= 113 && r <= 119 && cc == 18
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Purchase"] = value
                  end
                  if r >= 113 && r <= 119 && cc == 19
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = {}
                    @loan_amount["LoanAmount/LoanPurpose"][loan_key]["Refinance"] = value
                  end
                  if r == 119 && cc == 15 
                    @loan_amount["LoanSize/LoanPurpose"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"] = {}
                  end
                  if r == 119 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Purchase"] = value
                  end
                  if r == 119 && cc == 19
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["Conforming"]["Refinance"] = value
                  end
                  if r == 120 && cc == 18
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = {}
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Purchase"] =value
                    cc = cc + 1
                    new_val = sheet_data.cell(r,cc)
                    @loan_amount["LoanSize/LoanPurpose"]["High-Balance"]["Refinance"] = new_val
                  end
                  if r == 121 && cc == 19
                    @loan_amount["MiscAdjuster/LoanPurpose"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = {}
                    @loan_amount["MiscAdjuster/LoanPurpose"]["CA Escrow Waiver (Full or Taxes Only)"]["Refinance"] = value
                  end
                  if r == 122 && cc == 19
                    @loan_amount["PropertyType/LoanPurpose"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = {}
                    @loan_amount["PropertyType/LoanPurpose"]["Manufactured Home"]["Refinance"] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@property_hash,@sub_hash,@loan_amount]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_new_rez_wholesale5806_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  private

  def read_sheet
    file = File.join(Rails.root,  'OB_NewRez_Wholesale5806.xls')
    @xlsx = Roo::Spreadsheet.open(file)
  end

  def check_sheet_empty
    action =  params[:action]
    sheet_data = @xlsx.sheet(action) rescue @xlsx.sheet(action.upcase) rescue @xlsx.sheet(action.downcase) rescue @xlsx.sheet(action.capitalize)

    if sheet_data.first_row.blank?
      @msg = "Sheet is empty."
      redirect_to ob_new_rez_wholesale5806_index_path
    end
  end

  def get_sheet
    @sheet_obj = Sheet.find(params[:id])
  end

  def get_program
    @program = Program.find(params[:id])
  end

  def get_bank
    @bank = Bank.find(params[:id])
  end

  def get_titles
    return ["FICO/LTV Adjustments - Loan Amount ≤ $1MM", "State Adjustments", "FICO/LTV Adjustments - Loan Amount > $1MM", "Feature Adjustments", "Max Price"]
  end

  def all_lp
    data = Adjustment::ALL_IP

    return data
  end

  def high_bal_adjustment
    data = Adjustment::HIGH_BALANCE_ADJUSTMENT
    return data
  end

  def jumbo_series_i_adjustment
      data = Adjustment::JUMBO_SERIES_I_ADJUSTMENT
    return data
  end

  def dream_big_adjustment
    data = Adjustment::DREAM_BIG_ADJUSTMENT

    return data
  end

  def table_data
    hash_keys = {
      "FICO/LTV Adjustments" => "LoanAmount/FICO/LTV",
      "Feature Adjustments"  => "Feature/LTV",
      "State Adjustments" => "State",
      "Max Price" => "Max Price"
    }

    return hash_keys
  end


  # def make_adjust(block_hash, sheet)
  #   block_hash.keys.each do |key|
  #     unless ["Lender Paid MI Adj.", "Term/LTV/FICO"].include?(key)
  #       hash = {}
  #       hash[key] = block_hash[key]
  #       Adjustment.create(data: hash,loan_category: @sheet_name)
  #     else
  #       unless block_hash[key].empty?
  #         block_hash[key].keys.each do |s_key|
  #           h1 = {}
  #           h1[s_key] = block_hash[key][s_key]
  #           Adjustment.create(data: h1,loan_category: @sheet_name)
  #         end
  #       end
  #     end
  #   end
  # end

  def find_key(title)
    if title.eql?("FICO/LTV Adjustments - Loan Amount ≤ $1MM") or title.eql?("FICO/LTV Adjustments - Loan Amount > $1MM")
      base_key = table_data[@title.split(" -").first]
    else
      base_key = table_data[@title]
    end

    return base_key
  end

  def get_table_keys
    table_keys = Adjustment::MAIN_KEYS
    return table_keys
  end

  def get_value value1
    if value1.present?
      if value1.include?("<=") || value1.include?("<") || value1.include?("≤")
        value1 = "0-"+value1.split("<=").last.tr('A-Za-z%$><=≤, ','')
        value1 = value1.tr('–','-')
      elsif value1.include?(">") || value1.include?("+")
        value1 = value1.split(">").last.tr('A-Za-z+= ','')+"-Inf"
        value1 = value1.tr('–','-')
      elsif value1.include?("≥")
        value1 = value1.split("≥").last.tr('A-Za-z$, ','')+"-Inf"
        value1 = value1.tr('–','-')
      else
        value1 = value1.tr('$, ','')
        value1 = value1.tr('–','-')
      end
    end
  end

  def set_range value
    if value.split()[0].eql?("≤") || value.split()[0].eql?("<=") then
      value = "0-" + value.split()[1]
    elsif [">","≥",">=", "+"].include?(value.split()[0]) then
      value.split()[1] + "-#{Float::INFINITY}"
    elsif [">","≥",">=", "+"].include?(value.split("")[-1])
      value.split("+")[0] + "-#{Float::INFINITY}"
    elsif value.include?(">")
      value.split(">")[-1] + "-#{Float::INFINITY}"
    elsif value.include?("<=")
      value = "0-" + value.split("<=")[-1]
    end
  end

  def convert_range value
    if value.include?("≤") && value.include?("MM")
      value = "0-"+value.split("MM")[0].split("$")[-1] + "000000"
    elsif value.include?("-") && value.include?("MM") && value.include?("$")
      if value.split(" - ")[-1].split("MM")[0].split("$")[-1].include?(".")
        value = value.split(" - ")[0].split("MM")[0].split("$")[-1] + "000000" + "-" + value.split(" - ")[-1].split("MM")[0].split("$")[-1].gsub(".", "") + "000000"
      else
        value = value.split(" - ")[0].split("MM")[0].split("$")[-1].gsub(".", "") + "000000" + "-" + value.split(" - ")[-1].split("MM")[0].split("$")[-1] + "000000" if value.split(" - ")[0].split("MM")[0].split("$")[-1].include?(".")
      end
    end

    return value
  end

  def get_main_key heading
    heading.split(" ").each do |data|
      data.gsub!("#{data}", 'LoanType') if data.eql?("Fixed")
      data.gsub!("#{data}", 'Term') if data.eql?("terms")
    end
  end

  def program_property title
      @arm_caps = ''
      if title.downcase.exclude?("arm")
        term = title.downcase.split("fixed").first.tr('A-Za-z-™/® ','')
        if term.length == 4 && term.last(2).to_i < term.first(2).to_i
          term = term.last(2) + term.first(2)
        else
          term
        end
      end
         # Arm Basic
      if title.downcase.include?("arm")   
        if title.include?("3/1") || title.include?("3 / 1") || title.include?("3-1")
          arm_basic = 3
        elsif title.include?("5/1") || title.include?("5 / 1") || title.include?("5-1")
          arm_basic = 5
        elsif title.include?("7/1") || title.include?("7 / 1") || title.include?("7-1")
          arm_basic = 7
        elsif title.include?("10/1") || title.include?("10 / 1") || title.include?("10 /1") || title.include?("10-1")
          arm_basic = 10
        end
      end
      # Arm_advanced
      if title.downcase.include?("arm")
        title.split.each do |arm|
          if arm.tr('1-9A-Za-z()|.% ','') == "//" || arm.tr('1-9A-za-z() ','') == "--"
            @arm_caps = arm.tr('A-Za-z()|.% , ','')[0,5]
            if @arm_caps.include?('/')
              @arm_caps = @arm_caps.tr('/','-')
            else
              @arm_caps
            end
          end
        end
      end
      @program.update(term: term, arm_basic: arm_basic, arm_caps: @arm_caps)
    end

  def make_adjust(block_hash, sheet)
    block_hash.each do |hash|
      if hash.present?
        hash.each do |key|
          data = {}
          data[key[0]] = key[1]
          adj_ment = Adjustment.create(data: data,loan_category: sheet)
          link_adj_with_program(adj_ment, sheet)
        end
      end
    end
  end
end
