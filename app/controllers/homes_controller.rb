class HomesController < ApplicationController

  def banks
    @banks = Bank.all
  end
end
