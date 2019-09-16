class RenameBanksColumnName < ActiveRecord::Migration[5.2]
  def change
  	remove_column :banks, :state_eligibility
  	rename_column :banks, :state, :state_eligibility
  	add_column :banks, :state, :string
  end
end
