class DashboardController < SearchApi::DashboardController
  # before_action :profiling

  def home
    api_search
    @banks = Bank.all
    @all_banks_name = @banks.pluck(:name)
    @arm_advanced_list = @programs_all.pluck(:arm_advanced).push("5-5").compact.uniq.reject(&:empty?).map{|c| [c]}
    @arm_caps_list = @programs_all.pluck(:arm_caps).push("3-2-5").compact.uniq.reject(&:empty?).map{|c| [c]}
    @term_list = @programs_all.where('term <= ?', 999).pluck(:term).compact.uniq.push(5,10,15,20,25,30).uniq.sort.map{|y| [y.to_s + " yrs" , y]}.prepend(["All"])
    fetch_programs(true)
  end

  def fetch_programs(html_type=false)
    @programs_all = Program.all
    if params[:bank_name].present?
      @programs_all = @programs_all.where(bank_name: params[:bank_name]) unless params[:bank_name].eql?('All')
    end

    if params[:loan_category].present?
      @programs_all = @programs_all.where(loan_category: params[:loan_category]) unless params[:loan_category].eql?('All')
    end

    if params[:pro_category].present?
      @programs_all = @programs_all.where(program_category: params[:pro_category]) unless (params[:pro_category] == "All" || params[:pro_category] == "No Category")
    end

    @program_names = @programs_all.pluck(:program_name).uniq.compact.sort
    @loan_categories = @programs_all.pluck(:loan_category).uniq.compact.sort
    @program_categories = @programs_all.pluck(:program_category).uniq.compact.sort
    add_default_loan_cat
    render json: {program_list: @program_names.map{ |lc| {name: lc}}, loan_category_list: @loan_categories.map{ |lc| {name: lc}}, pro_category_list: @program_categories.map{ |lc| {name: lc}}} unless html_type
  end

  def set_ltv_value(ltv_amount)
    ltv = "65.00 - 69.99"
    Program::LTV_VALUES.each do |ltv_value|
      ltv_value = ltv_value[0]
      unless ltv_value.include?("+")
        lower = ltv_value.split("-").first.squish.to_f
        higher = ltv_value.split("-").last.squish.to_f
        if ltv_amount.between?(lower, higher)
          ltv = ltv_value
          break;
        end
      end
    end
    return ltv
  end

  def set_loan_amount(loan_amt)
    loan_amount = "250000 - 300000"
    if loan_amt <= 1550000
      Program::LOAN_AMOUNT.each do |loan_value|
        if loan_value.present? && loan_value.include?("-")
          lower = loan_value.split("-").first.squish.to_i
          higher = loan_value.split("-").last.squish.to_i
          if loan_amt.between?(lower, higher)
            loan_amount = loan_value
            break;
          end
        end
      end
    else
      loan_amount = "1550000 +"
    end
    return loan_amount
  end

  def set_la_and_ltv_value
    home_price = params[:home_price].present? ? params[:home_price].tr("^0-9.", '').to_f : 300000.00
    down_payment = params[:down_payment].present? ? params[:down_payment].tr("^0-9.", '').to_f : 50000.00
    loan_amt = home_price - down_payment
    loan_amount_range = set_loan_amount(loan_amt)
    ltv = loan_amt/home_price*100
    ltv_range  = set_ltv_value(ltv)
    respond_to do |format|
      format.json {render :json => {loan_amount_range: loan_amount_range , ltv_range: ltv_range, loan_amount: loan_amt, ltv: ltv }}
    end
  end

  def set_hv_and_dp_value
    loan_amt = params[:loan_amount_text].tr("^0-9.", '').to_f if params[:loan_amount_text].present?
    ltv = params[:ltv_text].tr("^0-9.", '').to_f if params[:ltv_text].present?
    loan_amount_range = set_loan_amount(loan_amt)
    ltv_range  = set_ltv_value(ltv)
    home_price =  (loan_amt/(ltv/100)).to_i
    down_payment = (home_price - loan_amt).to_i
    respond_to do |format|
      format.json {render :json => {loan_amount_range: loan_amount_range , ltv_range: ltv_range, home_price: home_price, down_payment: down_payment }}
    end
  end


  # def profiling
  #   Rack::MiniProfiler.authorize_request
  # end
end
