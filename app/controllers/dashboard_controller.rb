class DashboardController < ApplicationController
  before_action :set_default

  def index
    # @banks = Bank.all
    # if request.post?
    #   set_variable
    #   find_base_rate
    # end
  end

  def set_default
    @base_rate = 0.0
    @filter_data = {}

  end

  def set_variable
    if params[:rate_type].present?
      @filter_data[:rate_type] = params[:rate_type]

      if params[:rate_type] =="ARM" && params[:term_arm].present?
        @filter_data[:rate_arm] = params[:term_arm].to_i
      end

      if params[:rate_type] =="Fixed" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end

      if params[:rate_type] =="Floating" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end

      if  params[:rate_type] =="Variable" && params[:term].present?
        @filter_data[:term] = params[:term].to_i
      end
    end

    @interest = params[:interest] if params[:interest].present?
    @lock_period = params[:lock_period] if params[:lock_period].present?
        
    @gov_sheet = params[:gov] if params[:gov].present?
    if @gov_sheet.present?
      if @gov_sheet == "FHA"
        @filter_data[:fha] = true
      elsif @gov_sheet == "VA"
        @filter_data[:va] = true
      elsif @gov_sheet == "USDA"
        @filter_data[:usda] = true
      else
         @filter_data[:va] = false
         @filter_data[:fha] = false
         @filter_data[:usda] = false
      end
    end

    @loan_limit_type = params[:loan_limit_type] if params[:loan_limit_type].present?
    if @loan_limit_type.present?
      if @loan_limit_type == "Non-Conforming"
        @filter_data[:conforming] = false
      elsif @loan_limit_type == "Conforming"
        @filter_data[:conforming] = true
      elsif @loan_limit_type == "Jumbo"
        @filter_data[:Jumbo] = true
      elsif @loan_limit_type == "High-Balance"
        @filter_data[:jumbo_high_balance] = true
      else
        @filter_data[:conforming] = false
        @filter_data[:Jumbo] = false
        @filter_data[:jumbo_high_balance] = false
      end
    end

  end

  def find_base_rate
      @programs = Program.where(@filter_data)
      if @programs.present?
        if @programs.count>=2
          flash[:error] = "find multiple program"
        else
          program = @programs.first
          if program.base_rate[program.base_rate.keys.first][@interest.to_f.to_s].present?
            @base_rate = program.base_rate[program.base_rate.keys.first][@interest.to_f.to_s][@lock_period]
          else
            flash[:error] = "Not find any interest rate for this situation"
          end
        end
      else
        flash[:error] = "Not find any program for this situation"
      end
    return @base_rate
  end

  render :index
end
