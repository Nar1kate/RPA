import time
from calendar import monthrange

current_month = time.gmtime().tm_mon


def month():
    if current_month != 1:
        desired_month = current_month - 1
    else:
        desired_month = 12
    return monthrange(2024, desired_month)[1], desired_month


menu_1 = "site-header__menu-button"

menu_2 = "left-menu__mobile-header"



start_period = "//label[@for='fromDate']"
start_months = '//*[@id="x6yzHnS3PmDg26JWSJJHQUtgRdYhwntYBHnpGe6f7KXJI0S58gvlQbsdGrE0XLam"]/div[4]/div[1]/div[1]/div[1]'
change_to_start_month = f'/html/body/div[3]/div[4]/div/div/div[2]/div/div/div/div/div[5]/div[3]/div[2]/div[{month()[1]}]/div'
change_to_start_date = f'/html/body/div[3]/div[4]/div/div/div[2]/div/div/div/div/div[5]/div[3]/div[4]/div[3]/div[1]/div[contains(text(), 1)]'

end_period = "//label[@for='tillDate']"
end_months = '//*[@id="x6yzHnS3PmDg26JWSJJHQUtgRdYhwntYBHnpGe6f7KXJI0S58gvlQbsdGrE0XLam"]/div[7]/div[1]/div[1]/div[1]/div'
change_to_end_month = f'/html/body/div[3]/div[4]/div/div/div[2]/div/div/div/div/div[5]/div[3]/div[5]/div[{month()[1]}]'
change_to_end_date = f"/html/body/div[3]/div[4]/div/div/div[2]/div/div/div/div/div[5]/div[3]/div[7]/div[3]/div[5]/div[contains(text(), '{month()[0]}')]"

apply_period = '//span[contains(text(), "Показать")]'