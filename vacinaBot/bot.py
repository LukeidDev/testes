# Import for the Web Bot
from botcity.web import WebBot, Browser, By

# Import for integration with BotCity Maestro SDK
from botcity.maestro import *

# Disable errors if we are not connected to Maestro
BotMaestroSDK.RAISE_NOT_CONNECTED = False

import pandas as pd
from datetime import datetime


def main():
    # Runner passes the server url, the id of the task being executed,
    # the access token and the parameters that this task receives (when applicable).
    maestro = BotMaestroSDK.from_sys_args()
    ## Fetch the BotExecution with details from the task, including parameters
    execution = maestro.get_execution()

    print(f"Task ID is: {execution.task_id}")
    print(f"Task Parameters are: {execution.parameters}")

    bot = WebBot()

    # Configure whether or not to run on headless mode
    bot.headless = False

    # Uncomment to change the default Browser to Firefox
    bot.browser = Browser.FIREFOX

    # Uncomment to set the WebDriver path
    bot.driver_path = r"/home/lukeid/Documentos/geckodriver"

    # Opens the BotCity website.
    bot.browse("https://infoms.saude.gov.br/extensions/SEIDIGI_DEMAS_Vacina_C19/SEIDIGI_DEMAS_Vacina_C19.html")
    
    xpath_data = '/html/body/div[1]/div/div/div[5]/b[1]'
    xpath_doses = '/html/body/div[1]/div/div/div[6]/div[1]/div[1]/div[1]/div/div/article/div[1]/div/div/div/div/div/div[2]/div/div/div[1]/div/span'
    
    
    dom_data = bot.find_element(xpath_data, By.XPATH)
    dom_doses = bot.find_element(xpath_doses, By.XPATH)
    
    # data_str = dom_data.get_attribute('innerHTML')
    doses = dom_doses.get_attribute('innerHTML')
    data_str = '21/08/2023'
    
    #inicio do pandas
    data = datetime.strptime(data_str, '%d/%m/%Y')
    
    try:
        existing_df = pd.read_excel('teste.xlsx')
    except FileNotFoundError:
        existing_df = pd.DataFrame(columns=['coluna_data', 'coluna_dose'])
        
    data_format = data.strftime('%d/%m/%Y')
    
    if data_format in existing_df['coluna_data'].astype(str).values:
        print("Data já existe na planilha.")
    else:
        new_row = {'coluna_data': data_format, 'coluna_dose': doses}
        existing_df = existing_df._append(new_row, ignore_index=True)
        existing_df.to_excel("teste.xlsx", sheet_name="teste", index=False)
        print("Nova linha adicionada.")

    print(existing_df)
    

    # Implement here your logic...
    ...

    # Wait 3 seconds before closing
    bot.wait(3000)

    # Finish and clean up the Web Browser
    # You MUST invoke the stop_browser to avoid
    # leaving instances of the webdriver open
    bot.stop_browser()

    # Uncomment to mark this task as finished on BotMaestro
    # maestro.finish_task(
    #     task_id=execution.task_id,
    #     status=AutomationTaskFinishStatus.SUCCESS,
    #     message="Task Finished OK."
    # )


def not_found(label):
    print(f"Element not found: {label}")


if __name__ == '__main__':
    main()
