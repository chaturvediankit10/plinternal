class DashboardController < SearchApi::DashboardController
  # before_action :profiling

  def home
    api_search    
  end

  def fetch_programs
    fetch_programs_by_bank
  end

  # def profiling
  #   Rack::MiniProfiler.authorize_request
  # end
end
