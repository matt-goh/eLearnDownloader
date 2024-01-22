from tkinter import filedialog
from pyppeteer import launch
from tkinter import *
from tkinter.ttk import *
import tkinter as tk
import configparser
import requests
import asyncio
import sys
import re
import os

class Downloader:
    def __init__(self, app):
        self.app = app
        self.app.title("eLearn Downloader")
        self.app.iconbitmap('downloadericon.ico')
        # Makes the window unresizable
        self.app.resizable(0, 0)
        
        # Declaring string variables for entry widget
        self.url = StringVar()
        self.path = StringVar()
        self.userID = StringVar()
        self.userPassword = StringVar()
        
        # Load credentials if exists
        self.load_credentials()
        
        # Button Style
        style = tk.ttk.Style()
        style.configure('W.TButton', font=('calibri', 12))

        # Creating and defining widgets
        self.urlLabel = Label       (app, text='URL:', font=('calibri', 16, 'bold'))
        self.urlEntry = Entry       (app, textvariable=self.url, font=('calibri', 16), width=50)
        self.status = Label         (app, text='Status: Enter the link address of any eLearn page with download links.', font=('calibri', 12))
        self.filePath = Entry       (app, textvariable=self.path, font=('calibri', 12), width=50)
        self.changeBtn = Button     (app, text='Change', style='W.TButton', command=self.openFile)
        self.downloadBtn = Button   (app, text='Download', style='W.TButton', command=self.run)
        self.idLabel = Label        (app, text='Student ID:', font=('calibri', 12))
        self.id = Entry             (app, textvariable=self.userID, font=('calibri', 12), width=20)
        self.passLabel = Label      (app, text='Password:', font=('calibri', 12))
        self.password = Entry       (app, textvariable=self.userPassword, show='*', font=('calibri', 12), width=20)
        
        # Position all widgets using grid
        self.urlLabel.grid      (row=0, column=0, padx=(10, 0), pady=10)
        self.urlEntry.grid      (row=0, column=1, columnspan=2, padx=10, pady=10)
        self.status.grid        (row=1, column=0, columnspan=3)
        self.filePath.grid      (row=2, column=0, columnspan=3, pady=(10, 0))
        self.changeBtn.grid     (row=3, column=0, columnspan=3, padx=(0, 115), pady=10)
        self.downloadBtn.grid   (row=3, column=1, columnspan=3, padx=(45, 0), pady=10)
        self.idLabel.grid       (row=4, column=1, sticky=E)
        self.id.grid            (row=4, column=2, sticky=W, padx=(5, 65))
        self.passLabel.grid     (row=5, column=1, sticky=E)
        self.password.grid      (row=5, column=2, sticky=W, padx=(5, 65), pady=10)
                
        # self.testgrid = Button(app, text='here', style='W.TButton').grid(row=6, column=1, sticky=E)

    # Runs when user clicks the 'Change' button
    # Assign the new directory to 'newPath'
    def openFile(self):
        newPath = filedialog.askdirectory(initialdir=self.filePath)
        if newPath:
            self.filePath.delete(0, '')
            self.filePath.insert(0, newPath)
    
    # Runs when user clicks download button
    def run(self):
        
        current_url = self.url.get()
        userID = self.userID.get()
        userPass = self.userPassword.get()
        
        if not current_url:
            self.status.configure(foreground='black')
            self.status.config(text='Status: URL is empty.')
            return
        
        self.url.set("")
        
        self.save_credentials()
        
        asyncio.get_event_loop().run_until_complete(self.download(userID, userPass, current_url))
        
    async def download(self, userID, userPass, url):
        # Get the directory of the program executable
        basePath = getattr(sys, '_MEIPASS', os.path.abspath(os.path.dirname(__file__)))

        # Construct the path to Chromium
        chromiumPath = os.path.join(basePath, "chromium", "chrome.exe")

        # Launch Pyppeteer using the specified Chromium executable path
        browser = await launch(executablePath=chromiumPath, headless=False)
        
        page = await browser.newPage()
            
        await page.goto(url)
        await page.waitForNavigation(timeout=10000)

        try:
            # For the cookies policy consent form pop up that appears for new users
            # await page.waitForSelector('#agree_button')
            # await page.click('#agree_button')
            
            # Prints html contents for debugs
            # html_content = await page.content()
            # print(html_content)
            
            # Screenshot for debug purpose
            # screenshot_path = 'C:\\Users\\YourName\\Downloads\\screenshot.png'
            # await page.screenshot({'path': screenshot_path, 'type': 'png'})
            
            # Selects the sunway ID login button through the button's onClick function
            await page.evaluate('useADFS()')
            
            # Fill in the username and password fields
            await page.waitForSelector('#userNameInput')
            await page.type('#userNameInput', userID)
            await page.type('#passwordInput', userPass)

            # Submit the login form
            await page.evaluate('Login()')
            await page.click('#submitButton')
            
            try:
                # Wait until the page shows downloadable links
                await page.waitForSelector('a[href*=bbcswebdav]', {'timeout': 3000})
                
                # Wait for page to load
                await asyncio.sleep(3)
                
                # Extract the name from the title of the page, this will later be named for the folder of all files
                folderTitle = await page.evaluate('''() => {
                    const titleText = document.getElementById('pageTitleText');
                    return titleText.innerText.trim();
                }''')
                
                # Extract names and links from the page that contain 'bbcswebdav' in their href, 
                # which is a common keyword used by eLearn that is in every resources uploaded by lecturer
                
                linkNames = await page.evaluate('''() => {
                    const linkNames = [];
                    document.querySelectorAll('a[href*=bbcswebdav]').forEach(link => {
                        const span = link.querySelector('span');
                        linkNames.push(span ? span.textContent.trim() : link.textContent.trim());
                    });
                    return linkNames;
                }''');
                
                links = await page.evaluate('''() => {
                    const links = [];
                    document.querySelectorAll('a[href*=bbcswebdav]').forEach(link => {
                        links.push(link.href);
                    });
                    return links;
                }''')
                
                # Sanitize
                folderTitle = self.sanitizeName(folderTitle)
                
                try:
                    # Creating the download path, joining the path user entered with a new folder named with title of the page
                    downloadPath = os.path.join(self.filePath.get(), folderTitle)
                    os.makedirs(downloadPath, exist_ok=True)
                except:
                    self.status.configure(foreground='red')
                    self.status.config(text='Status: Download path error. Please make sure the path is valid.')
                
                # Capture cookies
                cookies = await page.cookies()
                
                # Debug purpose
                # print(cookies)
                # print(f"Titles: {folderTitle}")
                # print(f"Number of link names: {len(linkNames)}")
                # print(linkNames)
                # print(f"Number of links: {len(links)}")
                # print(links)
                
                # Run through all the names and links from two arrays and download them one by one
                for index in range(len(linkNames)):
                    self.downloadFile(linkNames[index], links[index], downloadPath, cookies)
                    
                self.status.configure(foreground='#4BB543')
                self.status.config(text='Status: Download Completed!')
            
            except Exception as page_error:
                self.status.configure(foreground='red')
                self.status.config(text=f'Status: Page error. Please make sure URL is correct.')
                print(page_error)

        except:
            self.status.configure(foreground='red')
            self.status.config(text='Status: Login failed.')
            
        finally:
            await browser.close()
            
    # Function for sanitizing
    def sanitizeName(self, name):
        invalidChars = re.compile(r'[\\/:"*?<>|]+') # Define a regular expression pattern to match invalid characters
        
        # Replace invalid characters with underscores
        sanitizedName = re.sub(invalidChars, '_', name)
            
        return sanitizedName
        
    def downloadFile(self, fileName, linkUrl, Dpath, cookies):
        fileName = self.sanitizeName(fileName)
        
        # Checks if the file contain extension or no, adds '.zip' if no
        _, file_extension = os.path.splitext(fileName)
        
        if not file_extension:
            fileName += '.zip'
        
        # Join the download path with the filename one last time
        finalFilePath = os.path.join(Dpath, fileName)

        # Download the file
        response = requests.get(linkUrl, cookies={cookie['name']: cookie['value'] for cookie in cookies})
        response.raise_for_status()
        
        # Open the final file path as a file, then write the content of the link into that file
        with open(finalFilePath, 'wb') as file:
            file.write(response.content)
            
    # Load credentials from credentials.ini
    def load_credentials(self):
        config = configparser.ConfigParser()
        try:
            config.read('credentials.ini')
            self.userID.set(config.get('Credentials', 'UserID'))
            self.userPassword.set(config.get('Credentials', 'UserPassword'))
            self.path.set(config.get('Credentials', 'path'))
        except configparser.Error:
            self.path.set(os.path.expanduser("~"))
            pass  # Handle the case where the config file or section does not exist

    # Save credentials to credentials.ini when user clicks download
    def save_credentials(self):
        config = configparser.ConfigParser()
        config['Credentials'] = {'UserID': self.userID.get(), 'UserPassword': self.userPassword.get(), 'path' : self.path.get()}

        with open('credentials.ini', 'w') as configfile:
            config.write(configfile)

app = tk.Tk()
Downloader(app)
app.mainloop()
