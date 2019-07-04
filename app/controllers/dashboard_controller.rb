class DashboardController < SearchApi::DashboardController
  # before_action :profiling

  def home
    list_of_banks_and_programs_with_search_results    
  end

  def fetch_programs
    fetch_programs_by_bank
  end

  def set_ltv_value(ltv_amount)
    ltv = "65.01 - 70.00"
    Program::LTV_VALUES.each do |ltv_value|
      ltv_value = ltv_value[0]
      unless ltv_value.include?("+")
        lower = ltv_value.split("-").first.squish.to_i
        higher = ltv_value.split("-").last.squish.to_i
        if ltv_amount.between?(lower, higher)
          ltv = ltv_value
        end
      end
    end
    return ltv
  end

  def set_loan_amount(loan_amt)
    loan_amount = "250000 - 300000"
    if loan_amt <= 850000
      Program::LOAN_AMOUNT.each do |a|
        loan_value = a[1]
        if loan_value.present? && loan_value.include?("-")
          lower = loan_value.split("-").first.squish.to_i
          higher = loan_value.split("-").last.squish.to_i
          if loan_amt.between?(lower, higher)
            loan_amount = loan_value
          end
        end
      end
    else
      loan_amount = "850000 +"
    end
    return loan_amount
  end

  def set_loan_amount_range
    home_price = params[:home_price].present? ? params[:home_price].tr("^0-9", '').to_f : 300000.00
    down_payment = params[:down_payment].present? ? params[:down_payment].tr("^0-9", '').to_f : 50000.00
    loan_amt = home_price - down_payment
    loan_amount_range = set_loan_amount(loan_amt)
    ltv_range  = set_ltv_value(loan_amt/home_price*100)
    respond_to do |format|
      format.json {render :json => {loan_amount: loan_amount_range , ltv: ltv_range}}
    end
  end


  # def profiling
  #   Rack::MiniProfiler.authorize_request
  # end
end
