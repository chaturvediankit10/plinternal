module ProgramAdj
  extend ActiveSupport::Concern
  $arr = ('A'..'ZZZ').to_a
	def link_adj_with_program(adj_ment, sheet)
		program_list = Program.where(loan_category: sheet)
		key_list = adj_ment.data.keys.first.split("/")
		program_filter1={}
		program_filter2={}
		include_in_input_values = false
		if key_list.present?
			key_list.each_with_index do |key_name, key_index|
			  if (Program.column_names.include?(key_name.underscore))
			    unless (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
			      program_filter1[key_name.underscore] = nil
			    else
			      if (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
			        program_filter2[key_name.underscore] = true
			      end
			    end
			    include_in_input_values = true
			  else
			    if(Adjustment::INPUT_VALUES.include?(key_name))
			      include_in_input_values = true
			    end
			  end
			end
			if (include_in_input_values)
			  program_list1 = program_list.where.not(program_filter1)
			  program_list2 = program_list1.where(program_filter2)
			  ids = ""
			  if program_list2.present?
			    program_list2.each do |program|
			      adj_id = program.adjustment_ids
			      if adj_id.present?
			        adj_id = adj_id+","+adj_ment.id.to_s
			      else
			        adj_id =adj_ment.id.to_s
			      end
			      program.update(:adjustment_ids => adj_id)
			    end
			  end
			end
		end
	end

	def get_cell_number main_key, row, col
		col = idxtoab col
    main_key["cell_number"] = row.to_s+""+col.to_s
    return
  end

  def idxtoab cc
  	$arr[cc-1]
  end

  def check_adjustment_range key
  	main_keys = ["LoanSize/LoanPurpose", "LoanType", "FICO", "VA/LoanAmount/LoanPurpose", "FHA/RefinanceOption/Streamline/VA", "LoanAmount/LoanPurpose", "LoanPurpose/LockDay", "FHA/LoanPurpose", "FHA/LoanType/VA/FICO", "LoanType/LoanPurpose/LockDay", "LoanType/VA/FICO", "VA/FICO", "VA/LoanPurpose/RefinanceOption", "FHA/VA/LoanSize/Term", "LoanSize/LoanType/Term/LTV/FICO", "RefinanceOption/LTV/FICO", "LPMI/PropertyType/FICO", "LPMI/Term/LTV/FICO", "PropertyType/LTV", "LPMI/RefinanceOption/FICO", "PropertyType", "PropertyType/Term/LTV", "MiscAdjuster", "FinancingType/LTV/CLTV/FICO", "LoanSize/LoanPurpose/RefinanceOption", "LoanSize/RefinanceOption", "MiscAdjuster/LoanPurpose", "PropertyType/LoanPurpose", "LoanSize/LoanType/LTV", "LoanSize/LoanType", "LoanSize/LoanType/LTV/FICO", "LPMI/LTV/FICO", "LTV", "LoanSize/LoanPurpose/RefinanceOption/LTV", "LoanSize/RefinanceOption/LTV", "LoanSize/FICO/LTV", "RefinanceOption/FICO/LTV", "LoanAmount/FICO/LTV", "LoanType/Term", "RefinanceOption/LTV", "MiscAdjuster/LTV", "State", "LoanPurpose/FICO/LTV", "LoanAmount/LTV", "LoanSize/LoanType/State/Term", "LoanSize/LoanType/State/ArmBasic", "LoanAmount", "LoanSize/LoanType/FICO/LTV", "LoanSize/LoanType/LoanAmount/LTV", "LoanSize/LoanType/PropertyType/LTV", "LoanSize/LoanType/LoanPurpose/Term/LTV", "LoanSize/LoanType/LoanPurpose/LTV", "LoanSize/LoanType/RefinanceOption/LTV", "LoanSize/LoanType/DTI/LTV", "LoanSize/LoanType/Term", "LoanSize/LoanType/ArmBasic", "MiscAdjuster/State", "LoanSize/LoanPurpose/LoanType/Term", "LoanSize/LoanPurpose/FICO/LTV", "LoanSize/LoanType/Term/LTV", "LoanSize/RefinanceOption/FICO/LTV", "LoanSize/LoanAmount/LTV", "LoanSize/LoanAmount/LoanType/Term", "LoanSize/LoanAmount/LoanType/ArmBasic", "LoanSize/State/LTV", "LoanSize/PropertyType/LTV", "LoanSize/MiscAdjuster/LTV", "LoanSize/LoanType/Term/FICO/LTV", "LoanSize/LoanType/RefinanceOption/Term/LTV", "LoanSize/LoanType/PropertyType/Term/LTV", "LoanType/ArmBasic", "FHA/USDA/FICO", "FHA/USDA/LoanPurpose", "FHA/USDA", "FHA/USDA/LoanSize/FICO", "FHA/Streamline/CLTV", "FHA/USDA/PropertyType", "FHA/USDA/LoanAmount/FICO", "FHA/USDA/State", "FICO/LTV", "FICO/Term", "FreddieMac/FICO", "LoanSize/FICO", "FannieMaeProduct/FreddieMacProduct/FICO/LTV", "VA/LoanPurpose/LTV", "VA/RefinanceOption/LoanAmount", "VA/LoanSize/FICO", "VA/RefinanceOption/LTV", "VA/LoanAmount/FICO", "Term/FICO/LTV", "FinancingType", "VA", "USDA", "FHA", "VA/RefinanceOption", "State/LTV", "LoanPurpose/LTV"]
		if main_keys.include?(key.first)
			return true
		end
  end
end
