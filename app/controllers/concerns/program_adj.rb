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
end
