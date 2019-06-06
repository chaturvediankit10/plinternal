class DashboardController < SearchApi::DashboardController
  # before_action :profiling

  def home
    list_of_banks_and_programs_with_search_results    
  end

  def fetch_programs
    fetch_programs_by_bank
  end

  # def profiling
  #   Rack::MiniProfiler.authorize_request
  # end
end
