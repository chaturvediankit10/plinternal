Rails.application.routes.draw do
  get 'ob_cmg_wholesales/index'
  root :to => "import_files#index"
  resources :import_files, only: [:index] do
    member do
      get :programs
      get :import_government_sheet
      get :import_freddie_fixed_rate
      get :import_conforming_fixed_rate
      get :home_possible
      get :conforming_arms
      get :lp_open_acces_arms
      get :lp_open_access_105
      get :lp_open_access
      get :du_refi_plus_arms
      get :du_refi_plus_fixed_rate_105
      get :du_refi_plus_fixed_rate
      get :dream_big
      get :high_balance_extra
      get :freddie_arms
      get :jumbo_series_d
      get :jumbo_series_f
      get :jumbo_series_h
      get :jumbo_series_i
      get :jumbo_series_jqm
      get :import_HomeReadyhb_sheet
      get :import_homereddy_sheet
    end
  end
  # For details on the DSL available within this file, see http://guides.rubyonrails.org/routing.html
end
