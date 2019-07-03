class DashboardController < SearchApi::DashboardController
  # before_action :profiling

  def home
    list_of_banks_and_programs_with_search_results    
  end

  def fetch_programs
    fetch_programs_by_bank
  end

  def set_loan_amount_range
    loan_amount = "0 - 50000"
    home_price = params[:home_price].present? ? params[:home_price].tr("^0-9", '').to_i : 300000
    down_payment = params[:down_payment].present? ? params[:down_payment].tr("^0-9", '').to_i : 50000
    loan_amt = home_price - down_payment
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
    respond_to do |format|
      format.json {render :json => {:response => loan_amount}}
    end
  end


  # def profiling
  #   Rack::MiniProfiler.authorize_request
  # end
end
