module ProgramAdj
  extend ActiveSupport::Concern

  # def create_program_association_with_adjustment(sheet)
  #   programs = Program.where(loan_category: sheet)
  #   adjustments = Adjustment.where(loan_category: sheet)
  #   adjustment_list = Adjustment.where(loan_category: sheet)
  #   program_list = Program.where(loan_category: sheet)
  #   adjustment_list.each_with_index do |adj_ment, index|
  #     key_list = adj_ment.data.keys.first.split("/")
  #     program_filter1={}
  #     program_filter2={}
  #     include_in_input_values = false
  #     if key_list.present?
  #       key_list.each_with_index do |key_name, key_index|
  #         if (Program.column_names.include?(key_name.underscore))
  #           unless (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
  #             program_filter1[key_name.underscore] = nil
  #           else
  #             if (Program.column_for_attribute(key_name.underscore).type.to_s == "boolean")
  #               program_filter2[key_name.underscore] = true
  #             end
  #           end
  #           include_in_input_values = true
  #         else
  #           if(Adjustment::INPUT_VALUES.include?(key_name))
  #             include_in_input_values = true
  #           end
  #         end
  #       end

  #       if (include_in_input_values)
  #         program_list1 = program_list.where.not(program_filter1)
  #         program_list2 = program_list1.where(program_filter2)
  #         ids = ""
  #         if program_list2.present?
  #           program_list2.each do |program|
  #             adj_id = ((ids.length > 0) ? "," + adj_ment.id.to_s : adj_ment.id.to_s) unless program.adjustment_ids.include?(adj_ment.id)
  #             ids << adj_id
  #             debugger
  #             program.update(:adjustment_ids => ids)
  #           end
  #         end
  #       end
  #     end
  #   end
  # end

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


			        # adj_id = ((ids.split(",").length > 0) ? "," + adj_ment.id.to_s : adj_ment.id.to_s) unless program.adjustment_ids.present? && program.adjustment_ids.include?(adj_ment.id.to_s)
			        # ids << adj_id.to_s
			        # debugger
			        # program.update(:adjustment_ids => ids)
			    end
			  end
			end
		end
	end
end