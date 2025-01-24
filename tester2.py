import pandas as pd
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
import time

# === Step 1: Read Excel data and filter rows =================================
excel_file = "C:/Users/Thais/Documents/Tester/Mappe1.xlsx"

df = pd.read_excel(excel_file)

filtered_df = df[
    (
        (df["Column 4"].isna()) | (df["Column 4"] == "")
    ) &
    (
        (df["Column 5"].isna()) | (df["Column 5"] == "")
    )
]

result_list = list(zip(filtered_df["Column 1"], filtered_df["Column 2"]))

print("Rows where columns 4 and 5 are empty:")
for item in result_list:
    print(item)

# === Step 2: Set up Selenium WebDriver =======================================
# EXAMPLE: using Chrome; adjust for the browser you prefer
driver = webdriver.Chrome()  # Make sure chromedriver is in PATH

# Navigate to Copilot login page
driver.get("https://m365.cloud.microsoft/")  # MODIFY_HERE_LOGIN_URL

# === Step 3: Perform login (PLACEHOLDER EXAMPLE) =============================
try:
    # 1) Press initial login button (if needed)
    login_field = driver.find_element(By.ID, "hero-banner-sign-back-in-to-office-365-link")  # e.g. "usernameInput"
    login_field.click()  # Replace with your username or handle securely

    # 1) Choose account
    choose_account = driver.find_element(By.Class, "table_row")  # e.g. "usernameInput"
    choose_account.click()  # Replace with your username or handle securely

    # 2) Find and fill the login/username field
    #username_field = driver.find_element(By.ID, "MODIFY_HERE_USERNAME_FIELD_ID")  # e.g. "usernameInput"
    #username_field.send_keys("YOUR_USERNAME")  # Replace with your username or handle securely

    # 3) Find and fill the password field
    password_field = driver.find_element(By.ID, "i0118")  # e.g. "passwordInput"
    password_field.send_keys("code")  # Replace with your password or handle securely

    # 4) Click the login/submit button
    login_button = driver.find_element(By.ID, "idSIButton9")  # e.g. "loginSubmit"
    login_button.click()

    # Wait for any post-login page load or redirect
    time.sleep(5)  # Adjust as necessary
except Exception as e:
    print("Login process encountered an error:", e)
    driver.quit()
    raise SystemExit()

# === Step 4: Loop through each pair and submit prompts =======================
# Example of storing responses in a list (each item might be the Copilot's answer)
responses = []

for col1_value, col2_value in result_list:
    try:
        # Craft an arbitrary prompt including the variables
        prompt = f"Hi Copilot, please do something with ID = {col1_value} and item = {col2_value}."

        # 1) Find the Copilot prompt input field (or text area)
        prompt_field = driver.find_element(By.ID, "MODIFY_HERE_PROMPT_TEXTAREA_ID")
        
        # Clear any existing text
        prompt_field.clear()
        
        # 2) Enter the prompt text
        prompt_field.send_keys(prompt)

        # 3) Click the button to submit the prompt
        submit_button = driver.find_element(By.ID, "MODIFY_HERE_SUBMIT_PROMPT_BUTTON_ID")
        submit_button.click()

        # 4) Wait for Copilot's response to appear (this will vary by site)
        #    Adjust sleep time or implement more sophisticated wait logic
        time.sleep(5)

        # 5) Locate the response text element
        #    (Might be in a <div> or <span> with a specific ID or class name)
        response_element = driver.find_element(By.ID, "MODIFY_HERE_RESPONSE_ELEMENT_ID")
        response_text = response_element.text

        # Append or print the response
        responses.append(response_text)
        print("Response for", col1_value, col2_value, ":\n", response_text)

    except Exception as e:
        print(f"Error occurred handling pair ({col1_value}, {col2_value}): {e}")
        # Decide whether to continue or break
        continue

# === Step 5: (Optional) Log out or close the browser =========================
# If you need to log out:
# logout_button = driver.find_element(By.ID, "MODIFY_HERE_LOGOUT_BUTTON_ID")
# logout_button.click()

# Close the browser
driver.quit()

# === Step 6: Do something with 'responses' if needed =========================
# e.g. write to a file or print summary
print("\nAll Copilot responses:")
for r in responses:
    print(r)
