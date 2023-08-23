from selenium import webdriver

# Set your IEDriverServer path
IEDRIVER_PATH = "/path/to/IEDriverServer.exe"

# Selenium configuration
capabilities = webdriver.DesiredCapabilities.INTERNETEXPLORER.copy()
capabilities["ignoreProtectedModeSettings"] = True
capabilities["nativeEvents"] = False

# Open browser and navigate to CITM
browser = webdriver.Ie(executable_path=IEDRIVER_PATH, capabilities=capabilities)
browser.get("CITM_URL")

print("hell")