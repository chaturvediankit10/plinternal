class ObQuickenLoans3571Controller < ApplicationController
  include ProgramAdj
  include BankState
  before_action :read_sheet, only: [:index,:ws_du_lp_pricing, :durp_lp_relief_pricing, :fha_usda_full_doc_pricing, :fha_streamline_pricing, :va_full_doc_pricing, :va_irrrl_pricing_govy_llpas, :na_jumbo_pricing_llpas, :du_lp_llpas, :durp_lp_relief_llpas, :lpmi]
	before_action :get_sheet, only: [:programs, :ws_du_lp_pricing, :durp_lp_relief_pricing, :fha_usda_full_doc_pricing, :fha_streamline_pricing, :va_full_doc_pricing, :va_irrrl_pricing_govy_llpas, :na_jumbo_pricing_llpas, :du_lp_llpas, :durp_lp_relief_llpas, :lpmi]
  before_action :get_program, only: [:single_program]

	def index
    begin
      @xlsx.sheets.each do |sheet|
        if (sheet == "WS Rate Sheet Summary")
          # headers = ["Phone", "General Contacts", "Mortgagee Clause (Wholesale)"]
          @name = "Quicken Loans"
          @bank = Bank.find_or_create_by(name: @name, state_eligibility: state_code_by_bank(@name))
          detail = get_bank_info(@name)
          @bank.update(address1: detail[:address1],address2:detail[:address2],phone: detail[:phone] ,zip: detail[:zip],city: detail[:city],state: detail[:state], state_code: detail[:state_code])
        end
        @sheet = @bank.sheets.find_or_create_by(name: sheet)
      end
    rescue
      # the required headers are not all present
    end
  end

  def ws_du_lp_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "WS DU & LP Pricing")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (15..129).each do |r|
          row = sheet_data.row(r)
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = sheet
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title

                  if @title == "30 Year Home Possible/Home Ready"
                    @program.update(term: 30)
                  end

  	              @block_hash = {}
  	              key = ''
  	              (1..25).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end

      if (sheet == "DU & LP LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @subordinate_hash = {}
        @property_hash = {}
        @cashout_hash = {}
        @other_adjustment = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key1 = ''
        secondary_key = ''
        ltv_key = ''
        ltv_key1 = ''
        cltv_key = ''
        new_key = ''
        # Adjustments
        (27..66).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(28)
          @cltv_data = sheet_data.row(38)
          @property_data = sheet_data.row(54)
          if row.compact.count
            (0..21).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "DU & LP LTV/FICO; Terms > 15 Years, Including ARMs"
                    @adjustment_hash["LoanType/Term/FICO/LTV"] = {}
                    @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"] = {}
                    @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["15-Inf"] = {}
                    @adjustment_hash["LoanType/FICO/LTV"] = {}
                    @adjustment_hash["LoanType/FICO/LTV"]["ARM"] = {}
                  end
                  if value == "Subordinate Financing"
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"] = {}
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Multiple Unit Property"
                    @property_hash["FreddieMac/PropertyType/LTV"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"] = {}
                  end
                  if value == "Home Possible & Home Ready Adjustment Caps"
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"] = {}
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"] = {}
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"] = {}
                  end
                  if value == "Cash Out"
                    @cashout_hash["RefinanceOption/FICO/LTV"] = {}
                    @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"] = {}
                  end
                  if value == "ARM Structure"
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"] = {}
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"] = {}
                  end
                  if value == "High Balance & ARMs"
                    @other_adjustment["LoanSize/LoanType/LTV"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  end

                  # DU & LP LTV/FICO; Terms > 15 Years, Including ARMs
                  if r >= 29 && r <= 35 && cc == 3
                    secondary_key = get_value value
                    if r == 29
                      get_cell_number @adjustment_hash["LoanType/Term/FICO/LTV"],r,cc
                      get_cell_number @adjustment_hash["LoanType/FICO/LTV"],r,cc
                    end
                    @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["15-Inf"][secondary_key] = {}
                    @adjustment_hash["LoanType/FICO/LTV"]["ARM"][secondary_key] = {}
                  end
                  if r >= 29 && r <= 35 && cc >= 5 && cc <= 21
                    ltv_key = get_value @ltv_data[cc-1]
                    @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["15-Inf"][secondary_key][ltv_key] = {}
                    @adjustment_hash["LoanType/Term/FICO/LTV"]["Fixed"]["15-Inf"][secondary_key][ltv_key] = value
                    @adjustment_hash["LoanType/FICO/LTV"]["ARM"][secondary_key][ltv_key] = {}
                    @adjustment_hash["LoanType/FICO/LTV"]["ARM"][secondary_key][ltv_key] = value
                  end
                  # Subordinate Financing
                  if r >= 39 && r <= 42 && cc == 3
                    if r == 39
                      get_cell_number @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"],r,cc
                    end
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"] = {}
                  end
                  if r >= 39 && r <= 42 && cc == 4
                    secondary_key = get_value value
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key] = {}
                  end
                  if r >= 39 && r <= 42 && cc == 5
                    cltv_key = get_value value
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key] = {}
                  end
                  if r >= 39 && r <= 42 && cc >= 7 && cc <= 9
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = value
                  end
                  if r >= 43 && r <= 46 && cc == 3
                    if r == 43
                      get_cell_number @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"],r,cc
                    end
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"] = {}
                  end
                  if r >= 43 && r <= 46 && cc == 4
                    secondary_key = get_value value
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key] = {}
                  end
                  if r >= 43 && r <= 46 && cc == 5
                    cltv_key = get_value value
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key] = {}
                  end
                  if r >= 43 && r <= 46 && cc >= 7 && cc <= 9
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = value
                  end
                  if r == 47 && cc == 4
                    secondary_key = get_value value
                    @subordinate_hash["FinancingType/LTV/FICO"] = {}
                    get_cell_number @subordinate_hash["FinancingType/LTV/FICO"],r,cc
                    @subordinate_hash["FinancingType/LTV/FICO"]["Subordinate Financing"] = {}
                    @subordinate_hash["FinancingType/LTV/FICO"]["Subordinate Financing"][secondary_key] = {}
                  end
                  if r == 47 && cc >= 7 && cc <= 9
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/LTV/FICO"]["Subordinate Financing"][secondary_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/LTV/FICO"]["Subordinate Financing"][secondary_key][ltv_key] = value
                  end
                  # ARM Structure
                  if r >= 49 && r <= 51 && cc == 13
                    secondary_key = value.split('/').first
                    if r == 49
                      get_cell_number @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"],r,cc
                    end
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key] = {}
                  end
                  if r >= 49 && r <= 51 && cc == 17
                    ltv_key = value.tr('/','-')
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key] = {}
                  end
                  if r >= 49 && r <= 51 && cc == 19
                    cltv_key = (value*100)
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key][cltv_key] = {}
                  end
                  if r >=49 && r <= 51 && cc == 21
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key][cltv_key] = value
                  end
                  if r == 55 && cc == 13
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"] = {}
                    get_cell_number @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"],r,cc
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"] = {}
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"]["true"] = {}
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"]["true"]["true"] = {}
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"]["true"]["true"]["0-Inf"] = {}
                  end
                  if r == 55 && cc >= 15 && cc <= 21
                    ltv_key = get_value @property_data[cc-1]
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"]["true"]["true"]["0-Inf"][ltv_key] = {}
                    @other_adjustment["PropertyType/FannieMae/FreddieMac/LTV"]["Investment Property"]["true"]["true"]["0-Inf"][ltv_key] = value
                  end
                  if r == 58 && cc == 13
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/FreddieMac/Term/LTV"],r,cc
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"]["true"]["true"] = {}
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"]["true"]["true"]["15-Inf"] = {}
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"]["true"]["true"]["15-Inf"]["75-Inf"] = {}
                    cc = cc + 8
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/FreddieMac/Term/LTV"]["true"]["true"]["15-Inf"]["75-Inf"] = new_value
                  end
                  if r == 62 && cc == 13
                    @other_adjustment["LoanSize/LoanType/RefinanceOption"] = {}
                    get_cell_number @other_adjustment["LoanSize/LoanType/RefinanceOption"],r,cc
                    @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"] = {}
                    @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"] = {}
                    @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"]["Cash Out"] = {}
                    cc = cc + 8
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanSize/LoanType/RefinanceOption"]["High-Balance"]["ARM"]["Cash Out"] = new_value
                  end
                  if r == 63 && cc == 13
                    @other_adjustment["LoanSize/LoanType/FannieMaeProduct"] = {}
                    get_cell_number @other_adjustment["LoanSize/LoanType/FannieMaeProduct"],r,cc
                    @other_adjustment["LoanSize/LoanType/FannieMaeProduct"]["High-Balance"] = {}
                    @other_adjustment["LoanSize/LoanType/FannieMaeProduct"]["High-Balance"]["ARM"] = {}
                    @other_adjustment["LoanSize/LoanType/FannieMaeProduct"]["High-Balance"]["ARM"]["HomeReady"] = {}
                    cc = cc + 8
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanSize/LoanType/FannieMaeProduct"]["High-Balance"]["ARM"]["HomeReady"] = new_value
                  end
                  if r >= 64 && r <= 65 && cc == 13
                    secondary_key = get_value value
                    if r == 64
                      get_cell_number @other_adjustment["LoanSize/LoanType/LTV"],r,cc
                    end
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][secondary_key] = {}
                    cc = cc + 8
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][secondary_key] = new_value
                  end
                  if r == 66 && cc == 13
                    @other_adjustment["LoanSize/LoanType/LTV"] = {}
                    get_cell_number @other_adjustment["LoanSize/LoanType/LTV"],r,cc
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["90-Inf"] = {}
                    cc = cc + 8
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"]["90-Inf"] = new_value
                  end
                  # Cash Out
                  if r >= 39 && r <= 45 && cc == 13
                    secondary_key1 = get_value value
                    if r == 39
                      get_cell_number @cashout_hash["RefinanceOption/FICO/LTV"],r,cc
                    end
                    @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][secondary_key1] = {}
                  end
                  if r >= 39 && r <= 45 && cc >= 15 && cc <= 21
                    ltv_key1 = get_value @cltv_data[cc-1]
                    @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][secondary_key1][ltv_key1] = {}
                    @cashout_hash["RefinanceOption/FICO/LTV"]["Cash Out"][secondary_key1][ltv_key1] = value
                  end
                  # Multiple Unit Property
                  if r == 50 && cc == 3
                    secondary_key = value.split("Property").first.try(:squish)
                    @property_hash["FannieMae/FreddieMac/PropertyType"] = {}
                    get_cell_number @property_hash["FannieMae/FreddieMac/PropertyType"],r,cc
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"] = {}
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"] = {}
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"][secondary_key] = {}
                    cc = cc + 6
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"][secondary_key] = new_value
                  end
                  if r == 51 && cc == 3
                    secondary_key = value.split("Property").first
                    @property_hash["FannieMae/PropertyType"] = {}
                    get_cell_number @property_hash["FannieMae/PropertyType"],r,cc
                    @property_hash["FannieMae/PropertyType"]["true"] = {}
                    @property_hash["FannieMae/PropertyType"]["true"][secondary_key] = {}
                    cc = cc + 6
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMae/PropertyType"]["true"][secondary_key] = new_value
                  end
                  if r >= 52 && r <= 54 && cc == 3
                    if value.include?("<")
                      secondary_key = "0-"+value.split('<').last.tr('()A-Z ','')
                    elsif value.include?(">")
                      secondary_key = value.split('>').last.tr('()A-Z ','')+"-Inf"
                    else
                      secondary_key = value.split('LTV').last.tr('()A-Z ','')
                    end
                    if r == 52
                      get_cell_number @property_hash["FreddieMac/PropertyType/LTV"],r,cc
                    end
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"][secondary_key] = {}
                    cc = cc + 6
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"][secondary_key] = new_value
                  end
                  if r >= 58 && r <= 59 && cc == 3
                    secondary_key = get_value value
                    if r == 58
                      get_cell_number @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"],r,cc
                    end
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"][secondary_key] = {}
                  end
                  if r >= 58 && r <= 59 && cc == 5
                    if value.downcase.include?('all')
                      ltv_key = "0-Inf"
                    else
                      ltv_key = get_value value
                    end
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"][secondary_key][ltv_key] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"][secondary_key][ltv_key] = new_value
                  end
                  if r == 60 && cc == 5
                    ltv_key = get_value value
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"][secondary_key][ltv_key] = {}
                    cc = cc + 4
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMaeProduct/FreddieMacProduct/LTV/FICO"]["HomeReady"]["Home Possible"][secondary_key][ltv_key] = new_value
                  end
                  if r == 65  && cc == 3
                    @property_hash["FreddieMac/LoanType"] = {}
                    get_cell_number @property_hash["FreddieMac/LoanType"],r,cc
                    @property_hash["FreddieMac/LoanType"]["true"] = {}
                    @property_hash["FreddieMac/LoanType"]["true"]["ARM"] = {}
                    cc = cc + 6
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FreddieMac/LoanType"]["true"]["ARM"] = new_value
                  end
                  if r == 66  && cc == 3
                    @property_hash["FreddieMac/PropertyType/LTV"] = {}
                    get_cell_number @property_hash["FreddieMac/PropertyType/LTV"],r,cc
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"]["75-Inf"] = {}
                    cc = cc + 6
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"]["75-Inf"] = new_value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@subordinate_hash,@property_hash,@cashout_hash,@other_adjustment]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def durp_lp_relief_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "DURP & LP Relief Pricing")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (14..124).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = sheet
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..15).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: rr, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end

      if (sheet == "DURP & LP Relief LLPAs")
        sheet_data = @xlsx.sheet(sheet)
        @adjustment_hash = {}
        @lp_adjustment = {}
        @subordinate_hash = {}
        @property_hash = {}
        @pricing_cap = {}
        @other_adjustment = {}
        @cashout_hash = {}
        primary_key = ''
        primary_key1 = ''
        secondary_key1 = ''
        secondary_key = ''
        ltv_key = ''
        ltv_key1 = ''
        cltv_key = ''
        new_key = ''
        # Adjustments
        (29..79).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(30)
          @cltv_data = sheet_data.row(50)
          @du_data = sheet_data.row(78)
          if row.compact.count
            (0..25).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "DURP LTV/FICO; Terms > 15 Years, Including ARMs"
                    @adjustment_hash["FannieMae/Term/FICO/LTV"] = {}
                    @adjustment_hash["FannieMae/Term/FICO/LTV"]["true"] = {}
                    @adjustment_hash["FannieMae/Term/FICO/LTV"]["true"]["15-Inf"] = {}
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"] = {}
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"]["true"] = {}
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"]["true"]["ARM"] = {}
                  end
                  if value == "LP Relief LTV/FICO; Terms > 15 Years, Including ARMs"
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"] = {}
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"]["true"] = {}
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"]["true"]["15-Inf"] = {}
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"] = {}
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"]["true"] = {}
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"]["true"]["ARM"] = {}
                  end
                  if value == "Subordinate Financing"
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"] = {}
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                  end
                  if value == "Multiple Unit Property"
                    @property_hash["FreddieMac/PropertyType/LTV"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"] = {}
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"] = {}
                  end
                  if value == "LP Relief Pricing Caps"
                    @pricing_cap["FreddieMac/PropertyType/LTV"] = {}
                    @pricing_cap["FreddieMac/PropertyType/LTV"]["true"] = {}
                  end
                  if value == "ARM Caps"
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"] = {}
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"] = {}
                  end
                  if value == "DURP Pricing Caps"
                    @other_adjustment["FannieMae/LoanType/Term/LTV"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"] = {}
                  end
                  if value == "High Balance & ARMs"
                    @other_adjustment["LoanSize/LoanType/LTV"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"] = {}
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"] = {}
                  end
                  if value == "High LTV"
                    @other_adjustment["FannieMae/FreddieMac/LTV"] = {}
                    @other_adjustment["FannieMae/FreddieMac/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/FreddieMac/LTV"]["true"]["true"] = {}
                    @other_adjustment["FannieMae/Term/LTV"] = {}
                    @other_adjustment["FannieMae/Term/LTV"]["true"] = {}
                  end

                  # DURP LTV/FICO; Terms > 15 Years, Including ARMs
                  if r >= 31 && r <= 37 && cc == 3
                    secondary_key = get_value value
                    if r == 31
                      get_cell_number @adjustment_hash["FannieMae/Term/FICO/LTV"],r,cc
                      get_cell_number @adjustment_hash["FannieMae/LoanType/FICO/LTV"],r,cc
                    end
                    @adjustment_hash["FannieMae/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key] = {}
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key] = {}
                  end
                  if r >= 31 && r <= 37 && cc >= 5 && cc <= 25
                    ltv_key = get_value @ltv_data[cc-1]
                    @adjustment_hash["FannieMae/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key][ltv_key] = {}
                    @adjustment_hash["FannieMae/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key][ltv_key] = value
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key][ltv_key] = {}
                    @adjustment_hash["FannieMae/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key][ltv_key] = value
                  end
                  # LP Relief LTV/FICO; Terms > 15 Years, Including ARMs
                  if r >= 41 && r <= 47 && cc == 3
                    secondary_key = get_value value
                    if r == 41
                      get_cell_number @lp_adjustment["FreddieMac/Term/FICO/LTV"],r,cc
                      get_cell_number @lp_adjustment["FreddieMac/LoanType/FICO/LTV"],r,cc
                    end
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key] = {}
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key] = {}
                  end
                  if r >= 41 && r <= 47 && cc >= 5 && cc <= 25
                    ltv_key = get_value @ltv_data[cc-1]
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key][ltv_key] = {}
                    @lp_adjustment["FreddieMac/Term/FICO/LTV"]["true"]["15-Inf"][secondary_key][ltv_key] = value
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key][ltv_key] = {}
                    @lp_adjustment["FreddieMac/LoanType/FICO/LTV"]["true"]["ARM"][secondary_key][ltv_key] = value
                  end
                  # DURP Pricing Caps
                  if r == 50 && cc == 16
                    @other_adjustment["FannieMae/LoanType/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/LoanType/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["Fixed"] = {}
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["ARM"] = {}
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["Fixed"]["0-80"] = {}
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["ARM"]["0-80"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["Fixed"]["0-80"] = new_value
                    @other_adjustment["FannieMae/LoanType/LTV"]["true"]["ARM"]["0-80"] = new_value
                  end
                  if r == 51 && cc == 18
                    get_cell_number @other_adjustment["FannieMae/LoanType/Term/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["0-20"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["0-20"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["0-20"]["80-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["0-20"]["80-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["0-20"]["80-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["0-20"]["80-Inf"] = new_value
                  end
                  if r == 52 && cc == 18
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["21-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["21-Inf"]["80-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["21-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["21-Inf"]["80-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["Fixed"]["21-Inf"]["80-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/Term/LTV"]["true"]["ARM"]["21-Inf"]["80-Inf"] = new_value
                  end
                  if r == 53 && cc == 18
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/LoanType/PropertyType/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["2nd Home"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["Investment Property"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["2nd Home"]["80-105"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["Investment Property"]["80-105"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["2nd Home"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["Investment Property"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["2nd Home"]["80-105"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["Investment Property"]["80-105"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["2nd Home"]["80-105"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["Fixed"]["Investment Property"]["80-105"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["2nd Home"]["80-105"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/LTV"]["true"]["ARM"]["Investment Property"]["80-105"] = new_value
                  end
                  if r == 54 && cc == 18
                    get_cell_number @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["26-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["26-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["26-Inf"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["26-Inf"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["26-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["26-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["26-Inf"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["26-Inf"]["105-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["26-Inf"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["26-Inf"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["26-Inf"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["26-Inf"]["105-Inf"] = new_value
                  end
                  if r == 55 && cc == 18
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["0-25"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["0-25"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["0-25"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["0-25"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["0-25"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["0-25"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["0-25"]["105-Inf"] = {}
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["0-25"]["105-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["2nd Home"]["0-25"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["Fixed"]["Investment Property"]["0-25"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["2nd Home"]["0-25"]["105-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanType/PropertyType/Term/LTV"]["true"]["ARM"]["Investment Property"]["0-25"]["105-Inf"] = new_value
                  end
                  if r == 56 && cc == 18
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/LoanSize/LoanType/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"] = {}
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"]["ARM"] = {}
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"]["ARM"]["0-80"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"]["ARM"]["0-80"] = new_value
                  end
                  if r == 57 && cc == 18
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"]["ARM"]["80-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanSize/LoanType/LTV"]["true"]["High-Balance"]["ARM"]["80-Inf"] = new_value
                  end
                  if r == 58 && cc == 18
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"],r,cc
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"] = {}
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["2nd Home"] = {}
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["Investment Property"] = {}
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["2nd Home"]["80-Inf"] = {}
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["Investment Property"]["80-Inf"] = {}
                    cc = cc + 7
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["2nd Home"]["80-Inf"] = new_value
                    @other_adjustment["FannieMae/LoanSize/PropertyType/LTV"]["true"]["High-Balance"]["Investment Property"]["80-Inf"] = new_value
                  end
                  if r >= 61 && r <= 63 && cc == 16
                    secondary_key = get_value value
                    if r == 61
                      get_cell_number @other_adjustment["LoanSize/LoanType/LTV"],r,cc
                    end
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][secondary_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["LoanSize/LoanType/LTV"]["High-Balance"]["ARM"][secondary_key] = new_value
                  end
                  if r == 64 && cc == 16
                    @other_adjustment["FreddieMac/LoanType"] = {}
                    get_cell_number @other_adjustment["FreddieMac/LoanType"],r,cc
                    @other_adjustment["FreddieMac/LoanType"]["true"] = {}
                    @other_adjustment["FreddieMac/LoanType"]["true"]["ARM"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FreddieMac/LoanType"]["true"]["ARM"] = new_value
                  end
                  # High LTV
                  if r >= 67 && r <= 68 && cc == 16
                    secondary_key = value.tr('A-Za-z()&% ','')
                    if r == 67
                      get_cell_number @other_adjustment["FannieMae/FreddieMac/LTV"],r,cc
                    end
                    @other_adjustment["FannieMae/FreddieMac/LTV"]["true"]["true"][secondary_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/FreddieMac/LTV"]["true"]["true"][secondary_key] = new_value
                  end
                  if r >= 69 && r <= 70 && cc == 16
                    secondary_key = value.split("DU").last.tr('A-Za-z) ','')
                    ltv_key = value.split("DU").first.tr('A-Za-z%(<> ','')+"-Inf"
                    if r == 69
                      get_cell_number @other_adjustment["FannieMae/Term/LTV"],r,cc
                    end
                    @other_adjustment["FannieMae/Term/LTV"]["true"][secondary_key] = {}
                    @other_adjustment["FannieMae/Term/LTV"]["true"][secondary_key][ltv_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FannieMae/Term/LTV"]["true"][secondary_key][ltv_key] = new_value
                  end
                  if r == 71 && cc == 16
                    @other_adjustment["FreddieMac/LTV"] = {}
                    get_cell_number @other_adjustment["FreddieMac/LTV"],r,cc
                    @other_adjustment["FreddieMac/LTV"]["true"] = {}
                    @other_adjustment["FreddieMac/LTV"]["true"]["105-Inf"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FreddieMac/LTV"]["true"]["105-Inf"] = new_value
                  end
                  if r == 74 && cc == 16
                    @other_adjustment["FreddieMac/PropertyType"] = {}
                    get_cell_number @other_adjustment["FreddieMac/PropertyType"],r,cc
                    @other_adjustment["FreddieMac/PropertyType"]["true"] = {}
                    @other_adjustment["FreddieMac/PropertyType"]["true"]["2nd Home"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["FreddieMac/PropertyType"]["true"]["2nd Home"] = new_value
                  end
                  if r == 75 && cc == 16
                    @other_adjustment["PropertyType/LTV"] = {}
                    get_cell_number @other_adjustment["PropertyType/LTV"],r,cc
                    @other_adjustment["PropertyType/LTV"]["Condo"] = {}
                    @other_adjustment["PropertyType/LTV"]["Condo"]["75-Inf"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @other_adjustment["PropertyType/LTV"]["Condo"]["75-Inf"] = new_value
                  end
                  if r == 79 && cc == 16
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"] = {}
                    get_cell_number @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"],r,cc
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"]["true"] = {}
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"]["true"]["true"] = {}
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"]["true"]["true"]["Investment Property"] = {}
                  end
                  if r == 79 && cc >= 18 && cc <= 25
                    ltv_key = get_value @du_data[cc-1]
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"]["true"]["true"]["Investment Property"][ltv_key] = {}
                    @other_adjustment["FannieMae/FreddieMac/PropertyType/LTV"]["true"]["true"]["Investment Property"][ltv_key] = value
                  end
                  # Subordinate Financing
                  if r >= 51 && r <= 53 && cc == 3
                    if r == 51
                      get_cell_number @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"],r,cc
                    end
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"] = {}
                  end
                  if r >= 51 && r <= 53 && cc == 5
                    secondary_key = get_value value
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key] = {}
                  end
                  if r >= 51 && r <= 53 && cc == 7
                    cltv_key = get_value value
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key] = {}
                  end
                  if r >= 51 && r <= 53 && cc >= 10 && cc <= 12
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/FannieMae/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = value
                  end
                  if r >= 54 && r <= 59 && cc == 3
                    if r == 54
                      get_cell_number @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"],r,cc
                    end
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"] = {}
                  end
                  if r >= 54 && r <= 59 && cc == 5
                    secondary_key = get_value value
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key] = {}
                  end
                  if r >= 54 && r <= 59 && cc == 7
                    cltv_key = get_value value
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key] = {}
                  end
                  if r >= 54 && r <= 59 && cc >= 10 && cc <= 12
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/FreddieMac/LTV/CLTV/FICO"]["Subordinate Financing"]["true"][secondary_key][cltv_key][ltv_key] = value
                  end
                  if r == 60 && cc == 5
                    if value.downcase.include?("all")
                      secondary_key = "0-Inf"
                    else
                      secondary_key = get_value value
                    end
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"] = {}
                    get_cell_number @subordinate_hash["FinancingType/LTV/CLTV/FICO"],r,cc
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"] = {}
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][secondary_key] = {}
                  end
                  if r == 60 && cc == 7
                    cltv_key = get_value value
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][secondary_key][cltv_key] = {}
                  end
                  if r == 60 && cc >= 10 && cc <= 12
                    ltv_key = get_value @cltv_data[cc-1]
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][secondary_key][cltv_key][ltv_key] = {}
                    @subordinate_hash["FinancingType/LTV/CLTV/FICO"]["Subordinate Financing"][secondary_key][cltv_key][ltv_key] = value
                  end
                  # Multiple Unit Property
                  if r == 63 && cc == 3
                    secondary_key = value.split("Property").first.try(:squish)
                    @property_hash["FannieMae/FreddieMac/PropertyType"] = {}
                    get_cell_number @property_hash["FannieMae/FreddieMac/PropertyType"],r,cc
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"] = {}
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"] = {}
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"][secondary_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMae/FreddieMac/PropertyType"]["true"]["true"][secondary_key] = new_value
                  end
                  if r == 64 && cc == 3
                    secondary_key = value.split("Property").first.try(:squish)
                    @property_hash["FannieMae/PropertyType"] = {}
                    get_cell_number @property_hash["FannieMae/PropertyType"],r,cc
                    @property_hash["FannieMae/PropertyType"]["true"] = {}
                    @property_hash["FannieMae/PropertyType"]["true"][secondary_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FannieMae/PropertyType"]["true"][secondary_key] = new_value
                  end
                  if r >= 65 && r <= 67 && cc == 3
                    if value.include?("<")
                      secondary_key = "0-"+value.split('<').last.tr('()A-Za-z% ','')
                    elsif value.include?(">")
                      secondary_key = value.split('>').last.tr('()A-Za-z% ','')+"-Inf"
                    else
                      secondary_key = value.split('LTV').last.tr('()A-Za-z% ','')
                    end
                    if r == 65
                      get_cell_number @property_hash["FreddieMac/PropertyType/LTV"],r,cc
                    end
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"][secondary_key] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @property_hash["FreddieMac/PropertyType/LTV"]["true"]["3-4 Unit"][secondary_key] = new_value
                  end
                  # LP Relief Pricing Caps
                  if r == 70 && cc == 3
                    get_cell_number @pricing_cap["FreddieMac/PropertyType/LTV"],r,cc
                    @pricing_cap["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"] = {}
                    @pricing_cap["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"]["0-80"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @pricing_cap["FreddieMac/PropertyType/LTV"]["true"]["2nd Home"]["0-80"] = new_value
                  end
                  if r == 71 && cc == 3
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"] = {}
                    get_cell_number @pricing_cap["FreddieMac/PropertyType/Term/LTV"],r,cc
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"] = {}
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"] = {}
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["0-20"] = {}
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["0-20"]["80-Inf"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["0-20"]["80-Inf"] = new_value
                  end
                  if r == 72 && cc == 3
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["21-Inf"] = {}
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["21-Inf"]["80-Inf"] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @pricing_cap["FreddieMac/PropertyType/Term/LTV"]["true"]["2nd Home"]["21-Inf"]["80-Inf"] = new_value
                  end
                  if r == 73 && cc == 3
                    @pricing_cap["FreddieMac/PropertyType"] = {}
                    get_cell_number @pricing_cap["FreddieMac/PropertyType"],r,cc
                    @pricing_cap["FreddieMac/PropertyType"]["true"] = {}
                    @pricing_cap["FreddieMac/PropertyType"]["true"][value] = {}
                    cc = cc + 9
                    new_value = sheet_data.cell(r,cc)
                    @pricing_cap["FreddieMac/PropertyType"]["true"][value] = new_value
                  end
                  # ARM Caps
                  if r >= 77 && r <= 79 && cc == 3
                    secondary_key = value.split('/').first
                    if r == 77
                      get_cell_number @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"],r,cc
                    end
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key] = {}
                  end
                  if r >= 77 && r <= 79 && cc == 5
                    ltv_key = value.tr('/','-')
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key] = {}
                  end
                  if r >= 77 && r <= 79 && cc == 8
                    cltv_key = (value*100)
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key][cltv_key] = {}
                  end
                  if r >= 77 && r <= 79 && cc == 11
                    @other_adjustment["LoanType/ArmBasic/ArmAdvanced/Margin"]["ARM"][secondary_key][ltv_key][cltv_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@lp_adjustment,@subordinate_hash,@property_hash,@pricing_cap,@other_adjustment]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def fha_usda_full_doc_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "FHA & USDA Full Doc Pricing")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..93).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = @sheet_name
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..20).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def fha_streamline_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "FHA Streamline Pricing")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..93).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = @sheet_name
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..20).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def va_full_doc_pricing
  	@xlsx.sheets.each do |sheet|
      if (sheet == "VA Full Doc Pricing")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        # programs
        (13..88).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = @sheet_name
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..20).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def va_irrrl_pricing_govy_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "VA IRRRL Pricing & Govy LLPAs")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @government_hash = {}
        primary_key1 = ''
        secondary_key = ''

        # programs
        (13..66).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = @sheet_name
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..20).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end

        (75..93).each do |r|
        	row = sheet_data.row(r)
        	(0..29).each do |cc|
        		begin
              value = sheet_data.cell(r,cc)
          		if value.present?
          			if value == "Government FICO Adjusters"
          				primary_key1 = "FICO"
          				@government_hash[primary_key1] = {}
          			end
                if value == "Loan Ladder"
                  @adjustment_hash["LoanAmount"] = {}
                end
          			if r == 76 && cc == 5
          				primary_key = "LockDay"
          				@adjustment_hash[primary_key] = {}
                  get_cell_number @adjustment_hash[primary_key],r,cc
                  @adjustment_hash[primary_key]["60"] = {}
        					cc = cc +10
        					new_value = sheet_data.cell(r,cc)
        					@adjustment_hash[primary_key]["60"] = new_value
          			end
          			# Loan Ladder
          			if r >= 88 && r <= 91 && cc == 5
          				primary_key = get_value value
                  if r == 88
                    get_cell_number @adjustment_hash["LoanAmount"],r,cc
                  end
          				@adjustment_hash["LoanAmount"][primary_key] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@adjustment_hash["LoanAmount"][primary_key] = new_value
          			end
          			# Government FICO Adjusters
          			if r >= 76 && r <= 79 && cc == 19
          				secondary_key = get_value value
                  if r == 76
                    get_cell_number @government_hash[primary_key1],r,cc
                  end
          				@government_hash[primary_key1][secondary_key] = {}
          				if @government_hash[primary_key1][secondary_key] = {}
          					cc = cc + 10
          					new_value = sheet_data.cell(r,cc)
          					@government_hash[primary_key1][secondary_key] = new_value
          				end
          			end
          			# Geography
          			if r == 82 && cc == 19
                  @government_hash["State"] = {}
                  get_cell_number @government_hash["State"],r,cc
                  @government_hash["State"]["NJ"] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash["State"]["NJ"] = new_value
          			end
          			if r == 83 && cc == 19
                  @government_hash["State"]["NY"] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash["State"]["NY"] = new_value
          			end
          			if r == 86 && cc == 19
          				@government_hash["VA/LoanPurpose/LTV"] = {}
                  get_cell_number @government_hash["VA/LoanPurpose/LTV"],r,cc
                  @government_hash["VA/LoanPurpose/LTV"]["true"] = {}
                  @government_hash["VA/LoanPurpose/LTV"]["true"]["Purchase"] = {}
                  @government_hash["VA/LoanPurpose/LTV"]["true"]["Purchase"]["95-Inf"] = {}

                  @government_hash["VA/RefinanceOption/LTV"] = {}
                  get_cell_number @government_hash["VA/RefinanceOption/LTV"],r,cc
                  @government_hash["VA/RefinanceOption/LTV"]["true"] = {}
                  @government_hash["VA/RefinanceOption/LTV"]["true"]["Rate and Term"] = {}
                  @government_hash["VA/RefinanceOption/LTV"]["true"]["Rate and Term"]["95-Inf"] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash["VA/LoanPurpose/LTV"]["true"]["Purchase"]["95-Inf"] = new_value
                  @government_hash["VA/RefinanceOption/LTV"]["true"]["Rate and Term"]["95-Inf"] = new_value
          			end
          			if r == 87 && cc == 19
                  @government_hash["VA/RefinanceOption/LTV"]["true"]["Cash Out"] = {}
                  @government_hash["VA/RefinanceOption/LTV"]["true"]["Cash Out"]["90-Inf"] = {}
        					cc = cc + 10
        					new_value = sheet_data.cell(r,cc)
        					@government_hash["VA/RefinanceOption/LTV"]["true"]["Cash Out"]["90-Inf"] = new_value
          			end
                if r == 89 && cc == 19
                  @government_hash["VA/LTV"] = {}
                  get_cell_number @government_hash["VA/LTV"],r,cc
                  @government_hash["VA/LTV"]["true"] = {}
                  @government_hash["VA/LTV"]["true"]["100-Inf"] = {}
                  cc = cc + 10
                  new_value = sheet_data.cell(r,cc)
                  @government_hash["VA/LTV"]["true"]["100-Inf"] = new_value
                end
          		end
            rescue Exception => e
              error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
              error_log.save
            end
        	end
        end
        adjustment = [@adjustment_hash,@government_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def na_jumbo_pricing_llpas
  	@xlsx.sheets.each do |sheet|
      if (sheet == "NA Jumbo Pricing & LLPAs")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        # programs
        (7..50).each do |r|
          row = sheet_data.row(r)
          row = row.reject { |e| e.to_s.empty? }
          if ((row.compact.count >= 1) && (row.compact.count <= 5))
            rr = r + 3
            max_column_section = row.compact.count - 1
            (0..max_column_section).each do |max_column|
              cc = 7*max_column + 5
              begin
                @title = sheet_data.cell(r,cc)
                if @title.present? && @title != "=FALSE()"
                  @program = @sheet_obj.programs.find_or_create_by(program_name: @title)
                  @program.loan_category = @sheet_name
                  Adjustment.where(loan_category: @program.loan_category).destroy_all
                  @programs_ids << @program.id
                  p_name = @title + " " + sheet
                  @program.update_fields p_name
                  program_property @title
  	              @block_hash = {}
  	              key = ''
  	              (1..15).each do |max_row|
  	                @data = []
  	                (0..3).each_with_index do |index, c_i|
  	                  rrr = rr + max_row
  	                  ccc = cc + c_i
  	                  value = sheet_data.cell(rrr,ccc)
  	                  if value.present?
  	                    if (c_i == 0)
  	                      key = value
  	                      @block_hash[key] = {}
  	                    else
  	                      @block_hash[key][15*c_i] = value if key.present?
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
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        # Adjustments
        (56..94).each do |r|
          row = sheet_data.row(r)
          if row.compact.count > 1
            (0..29).each do |cc|
              value = sheet_data.cell(r,cc)
              if value.present?
                if r == 70 && cc == 5
                  secondary_key = value.tr('A-Za-z><=, ','')+"-Inf"
                  @adjustment_hash["LoanAmount"] = {}
                  get_cell_number @adjustment_hash["LoanAmount"],r,cc
                  @adjustment_hash["LoanAmount"][secondary_key] = {}
                  cc = cc + 10
                  new_value = sheet_data.cell(r,cc)
                  @adjustment_hash["LoanAmount"][secondary_key] = new_value
                end
                if r == 73 && cc == 5
                  @adjustment_hash["PropertyType/Term/LTV"] = {}
                  get_cell_number @adjustment_hash["PropertyType/Term/LTV"],r,cc
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"] = {}
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["30"] = {}
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["30"]["0-60"] = {}
                  cc = cc + 10
                  new_value = sheet_data.cell(r,cc)
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["30"]["0-60"] = new_value
                end
                if r == 74 && cc == 5
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["15"] = {}
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["15"]["0-60"] = {}
                  cc = cc + 10
                  new_value = sheet_data.cell(r,cc)
                  @adjustment_hash["PropertyType/Term/LTV"]["Investment Property"]["15"]["0-60"] = new_value
                end
              end
            end
          end
        end
        adjustment = [@adjustment_hash]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def lpmi
    @xlsx.sheets.each do |sheet|
      if (sheet == "LPMI")
        @sheet_name = sheet
        sheet_data = @xlsx.sheet(sheet)
        @programs_ids = []
        @adjustment_hash = {}
        @additional_adjustment = {}
        secondary_key = ''
        ltv_key = ''
        # Adjustments
        (7..49).each do |r|
          row = sheet_data.row(r)
          @ltv_data = sheet_data.row(8)
          if row.compact.count >= 1
            (0..13).each do |cc|
              begin
                value = sheet_data.cell(r,cc)
                if value.present?
                  if value == "30-26 Years Fixed & ARMs"
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["Fixed"] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["Fixed"]["30-26"] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["ARM"] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["ARM"]["30-26"] = {}
                  end
                  if value == "30 Year Fixed: Freddie Home Possible "
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"] = {}
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"] = {}
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"]["Fixed"] = {}
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"]["Fixed"]["30"] = {}
                  end
                  if value == "25-21 Years Fixed"
                    @additional_adjustment["LoanType/Term/FICO"] = {}
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"] = {}
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["21-25"] = {}
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["8-20"] = {}
                    @additional_adjustment["PropertyType/FICO"] = {}
                  end
                  # 30-26 Years Fixed & ARMs
                  if r >= 9 && r <= 12 && cc == 3
                    secondary_key = get_value value
                    if r == 9
                      get_cell_number @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"],r,cc
                    end
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["Fixed"]["30-26"][secondary_key] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["ARM"]["30-26"][secondary_key] = {}
                  end
                  if r >= 9 && r <= 12 && cc >= 4 && cc <= 13
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["Fixed"]["30-26"][secondary_key][ltv_key] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["Fixed"]["30-26"][secondary_key][ltv_key] = value
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["ARM"]["30-26"][secondary_key][ltv_key] = {}
                    @adjustment_hash["LPMI/LoanType/Term/FICO/LTV"]["true"]["ARM"]["30-26"][secondary_key][ltv_key] = value
                  end
                  # 30 Year Fixed: Freddie Home Possible
                  if r >= 18 && r <= 21 && cc == 3
                    secondary_key = get_value value
                    get_cell_number @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"],r,cc
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"]["Fixed"]["30"][secondary_key] = {}
                  end
                  if r >= 18 && r <= 21 && cc >= 4 && cc <= 13
                    ltv_key = get_value @ltv_data[cc-2]
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"]["Fixed"]["30"][secondary_key][ltv_key] = {}
                    @adjustment_hash["FreddieMacProduct/LoanType/Term/FICO/LTV"]["Home Possible"]["Fixed"]["30"][secondary_key][ltv_key] = value
                  end
                  # 25-21 Years Fixed
                  if r >= 27 && r <= 30 && cc == 3
                    secondary_key = get_value value
                    if r == 27
                      get_cell_number @additional_adjustment["LoanType/Term/FICO"],r,cc
                    end
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["21-25"][secondary_key] = {}
                  end
                  if r >= 27 && r <= 30 && cc >= 4 && cc <= 13
                    ltv_key = get_value @ltv_data[cc-2]
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["21-25"][secondary_key][ltv_key] = {}
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["21-25"][secondary_key][ltv_key] = value
                  end
                  # 20-8 Years Fixed
                  if r >= 36 && r <= 39 && cc == 3
                    secondary_key = get_value value
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["8-20"][secondary_key] = {}
                  end
                  if r >= 36 && r <= 39 && cc >= 4 && cc <= 13
                    ltv_key = get_value @ltv_data[cc-2]
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["8-20"][secondary_key][ltv_key] = {}
                    @additional_adjustment["LoanType/Term/FICO"]["Fixed"]["8-20"][secondary_key][ltv_key] = value
                  end
                  # Misc Adjustments
                  if r >= 47 && r <= 49 && cc == 3
                    if value == "Inv. Prop"
                      secondary_key = "Investment Property"
                    else
                      secondary_key = get_value value
                    end
                    if r == 47
                      get_cell_number @additional_adjustment["PropertyType/FICO"],r,cc
                    end
                    @additional_adjustment["PropertyType/FICO"][secondary_key] = {}
                  end
                  if r >= 47 && r <= 49 && cc >= 4 && cc <= 13
                    ltv_key = get_value @ltv_data[cc-2]
                    @additional_adjustment["PropertyType/FICO"][secondary_key][ltv_key] = {}
                    @additional_adjustment["PropertyType/FICO"][secondary_key][ltv_key] = value
                  end
                end
              rescue Exception => e
                error_log = ErrorLog.new(details: e.backtrace_locations[0], row: r, column: cc, loan_category: sheet, error_detail: e.message)
                error_log.save
              end
            end
          end
        end
        adjustment = [@adjustment_hash,@additional_adjustment]
        make_adjust(adjustment,@sheet_name)
      end
    end
    redirect_to programs_ob_quicken_loans3571_path(@sheet_obj)
  end

  def programs
    @programs = @sheet_obj.programs
  end

  def single_program
  end

  def get_program
    @program = Program.find(params[:id])
  end

  private

    def get_sheet
      @sheet_obj = Sheet.find(params[:id])
    end

    def get_value value1
      if value1.present?
				if value1.include?("<=") || value1.include?("<")
          value1 = "0-"+value1.tr('A-Z<>=%$, ', '')
          value1 = value1.tr('–','-')
        elsif value1.include?(">=") || value1.include?(">") || value1.include?("+")
        	value1 = value1.tr('A-Z<>$%=+, ','')+"-Inf"
          value1 = value1.tr('–','-')
        elsif value1.include?("-")
          value1 = value1.tr('A-Za-z()/<>$%=, ','')
          value1 = value1.tr('–','-')
        else
          value1 = value1.tr('A-Za-z/()&%, ','')
          value1 = value1.tr('–','-')
        end
      end
    end

    def make_adjust(block_hash, sheet)
      block_hash.each do |hash|
        hash.each do |key|
          data = {}
          data[key[0]] = key[1]
          adj_ment = Adjustment.create(data: data,loan_category: sheet)
          link_adj_with_program(adj_ment, sheet)
        end
      end
    end

    def read_sheet
      file = File.join(Rails.root,  'OB_Quicken_Loans3571.xls')
      @xlsx = Roo::Spreadsheet.open(file)
    end

    def program_property title
      if (title.downcase.include?("year") || title.downcase.include?("yr") || title.downcase.include?("y")) && title.downcase.exclude?("arm")
        if title.scan(/\d+/).count > 1
          if(title.scan(/\d+/)[1].to_i < title.scan(/\d+/)[0].to_i)
            term = title.scan(/\d+/)[1] + term = title.scan(/\d+/)[0]
          else
            term = title.scan(/\d+/)[0] + term = title.scan(/\d+/)[1]
          end
        else
          term = title.scan(/\d+/)[0]
        end
      end
        # Arm Basic
      if title.downcase.include?("arm")
        arm_basic = title.scan(/\d+/)[0]
      end
      @program.update(term: term,arm_basic: arm_basic)
    end
end
