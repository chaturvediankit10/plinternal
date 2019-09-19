module BankState
  extend ActiveSupport::Concern

  def state_code_by_bank(bank_name)
    case bank_name
      when "Allied Mortgage"
        return %w[AL AR CA CO CT DE DC FL GA IL IN KS KY LA ME MD MA MI MN NH NJ NY NC OH OR PA RI SC TN TX VT VA WA WV AK AZ HI ID IA MS MO MT NE NV NM ND SD WI WY]

      when "CMG Financial"
        return %w[AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY OK UT]

      when "Home Point"
        return %w[AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY]

      when "Newfi Wholesale"
        return %w[AZ CA CO FL HI IL MD MI MN NJ OH OR PA TX WA WI] 

      when "NewRez"
        return %w[AL AK AR AZ CA CO CT DC DE FL GA GU HI ID IL IN IA KS KY LA MA ME MD MI MN MO MS MT NE NH NJ NM NV NY NC ND OH OK OR PA PR RI SC SD TN TX UT VA VI VT WA WV WI WY]

      when "Quicken Loans"
        return %w[AR AZ CA CO IL KS ME MA MS NV NH NJ NY OH OR PA RI TX VA WA]

      when "SunWest Wholesale"
        return %w[AL AK AZ AR CA CO CT DE DC FL HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY VI]

      when "Union Home"
        return %w[AL AZ AR CA CO CT DE DC FL GA IL IN IA KS KY LA ME MD MA MI MN MS MO NE NV NH NJ NM NC OH OK OR PA RI SC TN TX VT VA WA WV WI]

      when "United Wholesale"
        return %w[AK AL AR AZ CA CO CT DC DE FL GA IA ID IL IN KS KY LA MA MD ME MI MN MO MS MT NC ND NE NH NJ NM NV NY OH OK OR PA RI SC SD TN TX UT VA VT WA WI WV WY]

      when "Cardinal Financial"
        return %w[AL CA CO CT DE DC FL GA IL IN KS KY LA ME MD MA MI MN NH NJ NY NC OH OR PA RI SC TN TX VT VA WA WV]
      else
        return []
    end
  end

  def get_bank_info(bank_name)
    case bank_name
      when "Allied Mortgage"
        detail =  {:address1=>"225 E. City Ave, Suite 102",
                  :address2=> "225 E. City Ave, Suite 102",
                  :phone=> "(877) 448-2745",
                  :state=> "Pennsylvania",
                  :state_code=> "PA",
                  :zip=> "19004",
                  :city=> "Bala Cynwyd"}
        
        when "Cardinal Financial"
          detail =  {:address1=>"3701 Arco Corporate Drive, Suite 200",
                    :address2=> "3701 Arco Corporate Drive, Suite 200",
                    :phone=> "(855) 561-4944",
                    :state=> "North Carolina",
                    :state_code=> "NC",
                    :zip=> "28273",
                    :city=> "Charlotte"}
        
        when "CMG Financial"
          detail =  {:address1=>"3160 Crow Canyon Road Suite 400",
                    :address2=> "3160 Crow Canyon Road Suite 400",
                    :phone=> "(866) 659-8989",
                    :state=> "California",
                    :state_code=> "CA",
                    :zip=> "94583",
                    :city=> "San Ramon"}
        
        when "Home Point"
          detail =  {:address1=>"2211 Old Earhart Road, Suite 250",
                    :address2=> "2211 Old Earhart Road, Suite 250",
                    :phone=> "(800) 686-2404",
                    :state=> "Michigan",
                    :state_code=> "MI",
                    :zip=> "48105",
                    :city=> "Ann Arbor"}
        
        when "NewRez"
          detail =  {:address1=>"4000 Chemical Road, Suite 200",
                    :address2=> "4000 Chemical Road, Suite 200",
                    :phone=> "(866) 886-9285",
                    :state=> "Pennsylvania",
                    :state_code=> "PA",
                    :zip=> "19462",
                    :city=> "Plymouth Meeting"}
        
        when "Newfi Wholesale"
          detail =  {:address1=>"2200 Powell St, Suite 340",
                    :address2=> "2200 Powell St, Suite 340",
                    :phone=> "(888) 316-3934",
                    :state=> "California",
                    :state_code=> "CA",
                    :zip=> "94608",
                    :city=> "Emeryville"}
        
        when "Quicken Loans"
          detail =  {:address1=>"1050 Woodward Avenue",
                    :address2=> "1050 Woodward Avenue",
                    :phone=> "(877) 999-3811",
                    :state=> "Michigan",
                    :state_code=> "MI",
                    :zip=> "48226",
                    :city=> "Detroit"}
        
        when "SunWest Wholesale"
          detail =  {:address1=>"6131 Orangethorpe Ave Ste 500",
                    :address2=> "6131 Orangethorpe Ave Ste 500",
                    :phone=> "(844) 978-6937",
                    :state=> "California",
                    :state_code=> "CA",
                    :zip=> "90620",
                    :city=> "Buena Park"}
        
        when "Union Home"
          detail =  {:address1=>"8241 Dow Circle West",
                    :address2=> "8241 Dow Circle West",
                    :state=> "Ohio",
                    :state_code=> "OH",
                    :zip=> "44136",
                    :city=> "Strongsville"}
        
        when "United Wholesale"
          detail =  {:address1=>"585 South Blvd E.",
                    :address2=> "585 South Blvd E.",
                    :phone=> "(800) 981-8898",
                    :state=> "Michigan",
                    :state_code=> "MI",
                    :zip=> "48341",
                    :city=> "Pontiac"}
      end
    end
end