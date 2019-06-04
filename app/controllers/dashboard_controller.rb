class DashboardController < SearchApi::DashboardController

  def home
    list_of_banks_and_programs_with_search_results    
  end

  def fetch_programs
    fetch_programs_by_bank
  end
end
