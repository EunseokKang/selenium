
    actions = ActionChains(dr)
    actions.send_keys(Keys.ENTER).perform()
    element = wait.until(EC.presence_of_element_located((By.ID, "divSttn")))

    if a == '서구':
        gu = dr.find_element(By.ID, "label_30170")
        gu.click()
    elif a == '중구':