from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

def get_Results():
    output = 0
    driver = None
    try:
        options = webdriver.ChromeOptions()
        options.add_argument("--headless")
        options.add_argument("--no-sandbox")
        options.add_argument("--disable-dev-shm-usage")
        options.add_argument("--window-size=1920,1080")

        driver = webdriver.Chrome(options=options)
        driver.get("https://formulax.streamlit.app/")

        wait = WebDriverWait(driver, 20)

        wait.until(EC.invisibility_of_element_located(
            (By.CSS_SELECTOR, "[data-testid='stStatusWidget']")
        ))

        generate_btn = wait.until(EC.element_to_be_clickable(
            (By.XPATH, "//span[.//p[normalize-space(text())='Generate Report']]")
        ))
        generate_btn.click()

    except Exception as e:
        output = e
    finally:
        if driver:
            driver.quit()

    return output


getResultsOutput = get_Results()
if getResultsOutput != 0:
    print(f"Error: {getResultsOutput}")
else:
    print("Success!")
