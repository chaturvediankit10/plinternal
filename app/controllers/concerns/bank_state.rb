module BankState
  extend ActiveSupport::Concern

  def state_code_by_bank(bank_name)
    case bank_name
      when "Allied Mortgage"
        return %w[AL CA CO CT DE DC FL GA IL IN KS KY LA ME MD MA MI MN NH NJ NY NC OH OR PA RI SC TN TX VT VA WA WV]

      when "CMG Financial"
        return %w[AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY]

      when "Home Point"
        return %w[AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY]

      when "Newfi Wholesale"
        return %w[AZ CA CO FL HI IL MD MI MN NJ OH OR PA TX WA WI] 

      when "NewRez"
        return %w[AL AK AR AZ CA CO CT DC DE FL GA GU HI ID IL IN IA KS KY LA MA ME MD MI MN MO MS MT NE NH NJ NM NV NY NC ND OH OK OR PA PR RI SC SD TN TX UT VA VI VT WA WV WI WY]

      when "Quicken Loans"
        return %w[AR AZ CA CO IL KS ME MA MS NV NH NJ NY OH OR PA RI TX VA WA]

      when "SunWest Wholesale"
        return %w[AL AK AZ AR CA CO CT DE DC FL HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY]

      when "Union Home"
        return %w[AL AZ AR CA CO CT DE DC FL GA IL IN IA KS KY LA ME MD MA MI MN MS MO NE NV NH NJ NM NC OH OK OR PA RI SC TN TX VT VA WA WV WI]

      when "United Wholesale"
        return %w[AK AL AR AZ CA CO CT DC DE FL GA IA HI ID IL IN KS KY LA MA MD ME MI MN MO MS MT NC ND NE NH NJ NM NV NY OH OK OR PA RI SC SD TN TX UT VA VT WA WI WV WY]

      when "Cardinal Financial"
        return %w[AL AK AZ AR CA CO CT DE DC FL GA HI ID IL IN IA KS KY LA ME MD MA MI MN MS MO MT NE NV NH NJ NM NY NC ND OH OK OR PA RI SC SD TN TX UT VT VA WA WV WI WY]
      else
        return []
    end
  end
end