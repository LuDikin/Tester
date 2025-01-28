import os
import time
import re
import pandas as pd
import pyperclip  # <-- for copying text from clipboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from selenium.webdriver.chrome.options import Options

####################################################
# UPDATED SCRIPT TO HANDLE VARIABLE FORMATS & NANs #
####################################################
# Key changes:
# 1) We now parse Copilot's response in two ways:
#    - Try to find 4 quoted strings ("..."), i.e. using the regex "(.*?)"
#    - If not found, fallback to line-based parsing (split by newlines).
# 2) If some lines are missing or user typed N/A, we'll store them anyway.
# 3) We handle potential 'nan' values by converting to str() for row matching.
#
# The goal is to fill columns [Vægt, Enhed, Climatiq factor, Climatic factor navn]
# with whatever lines we can parse, including 'N/A' or partial text.
#
# Adjust excel_file path as needed.

excel_file = "C:/Users/Lukas/Documents/GBE5/P5Solution/CompESGDATa1.xlsx"  # <-- Change this path if needed

df = pd.read_excel(excel_file)

# Fill row references with empty strings if NaN, so we can compare them as strings.
df["Eksternt varenr."] = df["Eksternt varenr."].fillna("").astype(str)
df["Beskrivelse"] = df["Beskrivelse"].fillna("").astype(str)
df["Leverandør"] = df["Leverandør"].fillna("").astype(str)

# === Filter rows where Vægt, Enhed, Climatiq factor, Climatic factor navn are empty ===
filtered_df = df[(df["Vægt"].isna() | (df["Vægt"] == "")) &
                 (df["Enhed"].isna() | (df["Enhed"] == "")) &
                 (df["Climatiq factor"].isna() | (df["Climatiq factor"] == "")) &
                 (df["Climatic factor navn"].isna() | (df["Climatic factor navn"] == ""))]

# We'll gather the columns we need for the prompt:
result_list = list(zip(
    filtered_df["Eksternt varenr."],
    filtered_df["Beskrivelse"],
    filtered_df["Leverandør"]
))

print("Rows where Vægt, Enhed, Climatiq factor, Climatic factor navn are empty:")
for item in result_list:
    print(item)

# === Step 2: Set up Selenium WebDriver ========================================
try:
    driver = webdriver.Chrome()
except Exception as e:
    print("Could not start Chrome WebDriver:", e)
    raise SystemExit()

# === Step 3: Log in to M365 (with 2FA) ========================================

driver.get("https://m365.cloud.microsoft/")
time.sleep(3)

try:
    login_button = driver.find_element(By.ID, "hero-banner-sign-in-microsoft-365-link")
    login_button.click()
    time.sleep(2)

    username = os.environ.get("COPILOT_USERNAME", "")
    username_field = driver.find_element(By.ID, "i0116")
    username_field.send_keys(username)

    next_button = driver.find_element(By.ID, "idSIButton9")
    next_button.click()
    time.sleep(2)

    password = os.environ.get("COPILOT_PASSWORD", "")
    password_field = driver.find_element(By.ID, "i0118")
    password_field.send_keys(password)
    time.sleep(2)

    next_button = driver.find_element(By.ID, "idSIButton9")
    next_button.click()
    time.sleep(2)

    send_code_button = driver.find_element(
        By.CSS_SELECTOR,
        'div.table[role="button"][data-value="OneWaySMS"]'
    )
    send_code_button.click()
    time.sleep(2)

    twofa_code = input("Please enter the 2FA code you received by SMS: ")
    code_field = driver.find_element(By.ID, "idTxtBx_SAOTCC_OTC")
    code_field.send_keys(twofa_code)
    time.sleep(1)

    verify_button = driver.find_element(By.ID, "idSubmit_SAOTCC_Continue")
    verify_button.click()
    time.sleep(2)

    # (Optional) "Stay signed in?"
    try:
        stay_signed_in_btn = driver.find_element(By.ID, "idSIButton9")
        stay_signed_in_btn.click()
    except:
        pass

    time.sleep(5)

    # Go directly to Copilot Chat
    driver.get("https://m365.cloud.microsoft/chat?auth=2")
    time.sleep(3)

except Exception as e:
    print("Login process encountered an error:", e)
    driver.quit()
    raise SystemExit()

# === Step 4: Prepare for chat & define a more robust parser ==================
responses = []

time.sleep(3)  # let the page load iframes
iframes = driver.find_elements(By.TAG_NAME, "iframe")
print(f"Found {len(iframes)} iframes. Switching to the first one by default.")

if len(iframes) == 0:
    print("No iframes found, cannot proceed.")
    driver.quit()
    raise SystemExit()

driver.switch_to.frame(iframes[0])

def parse_copilot_response(copied_text: str):
    """
    Attempt to parse 4 lines from the Copilot response.
    1) First try finding 4 quoted matches: "(.*?)".
    2) If not found, fallback to line-based parse.
    3) If lines < 4, pad with 'N/A'.
    """
    # 1) Try quoting
    matches = re.findall(r'"(.*?)"', copied_text)
    if len(matches) == 4:
        return matches[0], matches[1], matches[2], matches[3]

    # 2) If that fails, try line-based (split)
    lines = [ln.strip() for ln in copied_text.splitlines() if ln.strip()]
    # if fewer than 4 lines, pad them with 'N/A'
    while len(lines) < 4:
        lines.append("N/A")
    return lines[0], lines[1], lines[2], lines[3]

# === Step 5: Loop over each row & prompt Copilot =============================
for ekst_varenr, beskrivelse, leverandor in result_list:
    try:
        # We'll keep a short example prompt. Adjust as needed.
        prompt_text = (f"Hi Copilot, can you find the net weight of the component or product with the ID {ekst_varenr} from the supplier by the name {leverandor}:\n"
                       f"It might be a help or a good clue to know the description of the product which is: {beskrivelse}. This way you know what product or component to look for when you browse the web, and you should also get familiar with the product/component and its material composition to the extent that is possible, so that you know what Climatiq factor to apply.\n"
                       f"So besides finding the net weight and the unit applied, I would like you to browse the Climatiq data explorer on the following URL 'https://www.climatiq.io/data/explorer'. \n"
                       f"In the Climatiq data explorer, find the best suitable emissions factor to apply for calculating carbon emissions of product/component with ID {ekst_varenr} from {leverandor}. Please, do only use cradle-to-shelf factors based on the material or whatever attributes the product/component may have.\n " 
                       f"As for your response to me, please provide four values for this item, with the exact format below. Please be very thorough when looking for Climatiq factors as I know this can be tedious. Also, don't write any other text than the four values I need, even if you can't find the value you are looking for. In such a case you where you can't find the value you should either enter N/A for that value if you couldn't find a reasonable answer. However, if you can come up with a reasonable value based on some logical assumptions, you may write the figure along with an 'AS' written in parenthesis along with the value you assumed.):\n"
                       f"On a last note about the response you're about to provide, do not make any assumptions about the Climatiq factor names \"<Climatic factor navn>\", and by name I mean the name of the factor and the LCA activity. Examples of what to respond could be 'Steel, primary production, cradle-to-shelf' or 'Aluminium extrusion, cradle-to-shelf' or 'Steel production, stainless, primary, at plant/RER S'."
                       f"Please respond in the format of a list with four lines, and do not write the unit applied for the weight in line 1, write only the unit in line 2. However, in line 3 you write both the emissions factor and the unit together, and in line 4 you write the name of the factor and the LCA activity. So just four lines only containing the net weight(1), unit of measurement(2), appropriate Climatiq factor found (3), and the found Climatiq factor name, and other attributes(4).\n"
                       #f"Line 1: \"<Vægt>\"\n"
                       #f"Line 2: \"<Enhed>\"\n"
                       #f"Line 3: \"<Climatiq factor>\"\n"
                       #f"Line 4: \"<Climatic factor navn>\"\n"
                       )

        # 1) Locate chat input
        chat_input = driver.find_element(By.CSS_SELECTOR, "[data-lexical-editor='true']")
        chat_input.click()

        # 2) Insert text
        driver.execute_script(
            """
            let inputField = arguments[0];
            inputField.focus();
            document.execCommand('insertText', false, arguments[1]);
            inputField.dispatchEvent(new Event('input', { bubbles: true }));
            """,
            chat_input,
            prompt_text
        )
        time.sleep(5)

        # 3) Send the prompt
        chat_input.send_keys(Keys.ENTER)
        time.sleep(10)

        # 4) Try copying the response
        try:
            copy_button = driver.find_element(By.CSS_SELECTOR, "span[data-automationid='splitbuttonprimary']")
            copy_button.click()
            time.sleep(1)

            copied_text = pyperclip.paste()
            print("Copied text:", copied_text)

            responses.append(f"Response for {ekst_varenr}, {beskrivelse}: {copied_text}")

            vaegt_val, enhed_val, climatiq_val, climatic_navn = parse_copilot_response(copied_text)

            # We store them even if they are 'N/A' or partial.
            row_mask = (
                (df["Eksternt varenr."].astype(str) == str(ekst_varenr)) &
                (df["Beskrivelse"].astype(str) == str(beskrivelse)) &
                (df["Leverandør"].astype(str) == str(leverandor))
            )

            if row_mask.any():
                df.loc[row_mask, "Vægt"] = vaegt_val
                df.loc[row_mask, "Enhed"] = enhed_val
                df.loc[row_mask, "Climatiq factor"] = climatiq_val
                df.loc[row_mask, "Climatic factor navn"] = climatic_navn
            else:
                print(f"Warning: No matching row found for {ekst_varenr}, {beskrivelse}, {leverandor}.")

        except Exception as copy_err:
            print("No copy button found or error copying:", copy_err)
            responses.append(f"No copy button found for ekst_varenr={ekst_varenr}")

        # (Optional) Start a new chat for the next item
        try:
            new_chat_btn = driver.find_element(By.CSS_SELECTOR, "button[data-automation-id='newChatButton']")
            new_chat_btn.click()
            time.sleep(3)
        except Exception as new_chat_err:
            print("Could not click newChatButton:", new_chat_err)

        print(f"Done prompt for ({ekst_varenr}, {beskrivelse}).\n")

    except Exception as e:
        print(f"Error occurred for ({ekst_varenr}, {beskrivelse}):", e)
        continue

# === Step 6: Close browser and print responses ===============================
driver.quit()

print("\nAll Copilot responses:")
for r in responses:
    print(r)

# === Step 7: Save updated DataFrame back to Excel ============================
df.to_excel(excel_file, index=False)
print("\nExcel updated with the 4 new columns (including N/A where relevant)!")
