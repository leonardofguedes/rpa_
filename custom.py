# from RPA.core.webdriver import download, start
import os
import re
import requests
import time
import logging
from datetime import datetime
from RPA.Excel.Files import Files
from RPA.Browser.Selenium import Selenium
from selenium.webdriver.common.by import By
from robocorp.tasks import get_output_dir
from datetime import datetime, timedelta
from selenium.common.exceptions import NoSuchElementException
from dateutil.relativedelta import relativedelta
from selenium.webdriver.chrome.options import Options

def create_headless_options():
    options = Options()
    options.add_argument("--headless")
    options.add_argument("--window-size=1920,1080")
    options.add_argument("--disable-gpu")
    options.add_argument("--no-sandbox")
    options.add_argument("--disable-dev-shm-usage")
    
    return options


class News:
    def __init__(self, title, link, source, time, description, image_url):
        self.title = title
        self.link = link
        self.source = source
        self.time = time
        self.description = description
        self.image_url = image_url

    def __repr__(self):
        return f"News(title={self.title}, link={self.link}, source={self.source}, time={self.time}, description={self.description}, image_url={self.image_url})"


class CustomSelenium:
    def __init__(self):
        self.driver = None
        self.logger = logging.getLogger(__name__)
        self.browser = Selenium(auto_close=False)
        self.articles = []
        self.pictures_dir = os.path.join(get_output_dir(), 'pictures')
        if not os.path.exists(self.pictures_dir):
            os.makedirs(self.pictures_dir)
        
    def ensure_directory_permissions(self, path):
        """Ensures the specified directory is writable."""
        if not os.access(path, os.W_OK):
            self.logger.error(f"Directory {path} is not writable.")
            raise PermissionError(f"Directory {path} is not writable.")

    @staticmethod
    def download_image(url, save_path):
        """
        Downloads an image from the specified URL and saves it to the given path.

        :param url: URL of the image to be downloaded.
        :param save_path: Local path to save the downloaded image.
        :return: True if the image was downloaded successfully, False otherwise.
        """
        try:
            response = requests.get(url, stream=True)
            if response.status_code == 200:
                with open(save_path, 'wb') as file:
                    for chunk in response.iter_content(1024):
                        file.write(chunk)
                return True
            else:
                return False
        except Exception as e:
            print(f"Error downloading image: {e}")
            return False
    
    def ensure_directory_exists(self, path):
        """Ensures the specified directory exists, creates it if not."""
        try:
            if not os.path.exists(path):
                os.makedirs(path)
            return True
        except OSError as e:
            self.logger.error(f"Error creating directory {path}: {e}")
            return False

    def find_and_click_news_link(self):
        """Finds and clicks the 'News' link on the page.

        This method waits until the 'News' link is visible and clickable,
        then clicks it. It will attempt multiple times in case of failure.
        """
        xpath = 'xpath://body//a[contains(@class, "d-ib") and contains(text(),"News")]'
        attempts = 3  # Number of attempts to try finding and clicking the element

        for attempt in range(1, attempts + 1):
            try:
                self.browser.wait_until_element_is_visible(
                    xpath, timeout='30s', error='Element "News" not visible after 30 seconds.'
                )
                self.browser.wait_until_element_is_enabled(
                    xpath, timeout='30s', error='Element "News" not clickable after 30 seconds.'
                )
                self.logger.info('Element "News" found and clickable. Clicking on the element.')
                self.browser.click_element(xpath)
                return  # Exit the function if successful
            except Exception as e:
                self.logger.error(f"Attempt {attempt} - Failed to find and click 'News' element: {e}")
                if attempt < attempts:
                    self.logger.info(f"Retrying... ({attempt}/{attempts})")
                    time.sleep(5)  # Wait a bit before retrying
                else:
                    self.logger.error("Exceeded maximum attempts to find and click 'News' element.")
                    raise  # Re-raise the last exception
    
    def filter_articles_by_date(self, months):
        """Filters articles based on the specified number of months.

        Args:
            months (int): The number of months to include in the filter.

        This method filters the self.articles list to only include articles from the specified
        number of months.
        """
        cutoff_date = datetime.now() - relativedelta(months=months)
        filtered_articles = [article for article in self.articles if self.relative_time_to_absolute(article.time) >= cutoff_date]
        self.articles = filtered_articles

    def collect_articles(self):
        """Collects articles from the Yahoo News search results page.

        This method locates and extracts information from news articles on the search results page,
        including the title, link, source, time, description, and image.

        The extracted data is stored in the self.articles list.

        Raises:
            NoSuchElementException: If an expected element is not found on the page.
        """
        articles_xpath = 'xpath://ol[@class="mb-15 reg searchCenterMiddle"]//li//div[@class="dd NewsArticle"]'
        elements = self.browser.get_webelements(articles_xpath)

        if not elements:
            self.logger.info("No articles found.")
            return

        for element in elements:
            try:
                article = self.extract_article_data(element)
                self.articles.append(article)
            except Exception as e:
                self.logger.error(f"Error extracting article data: {e}")

    def retry_action(self, func, retries=3, delay=2):
        """
        Attempts to execute the provided function multiple times with a delay between attempts.

        Args:
            func (function): The function to attempt.
            retries (int): Number of retries in case of failure.
            delay (int): Delay in seconds between attempts.

        Returns:
            The result of the function if successful.

        Raises:
            Exception: If the function fails after the specified number of retries.
        """
        start_time = time.time()
        for attempt in range(1, retries + 1):
            try:
                self.logger.info(f"Attempt {attempt} of {retries} for function {func.__name__}.")
                result = func()
                end_time = time.time()
                self.logger.info(f"Attempt {attempt} successful after {end_time - start_time:.2f} seconds.")
                return result
            except Exception as e:
                self.logger.error(f"Error on attempt {attempt} for function {func.__name__}: {e}")
                if attempt < retries:
                    self.logger.info(f"Retrying after {delay} seconds...")
                    time.sleep(delay)
                else:
                    self.logger.error(f"All {retries} attempts failed for function {func.__name__}.")
                    raise Exception(f"Action {func.__name__} failed after {retries} retries.") from e
        end_time = time.time()
        self.logger.info(f"Total time spent on retries: {end_time - start_time:.2f} seconds.")

    @staticmethod
    def contains_money(text):
        """Checks if the provided text contains any amount of money.

        This method uses regular expressions to identify different formats of monetary values
        within the provided text. Supported formats include:
        - $111,111.11
        - $11.1
        - 11 dollars
        - 11 USD

        Args:
            text (str): The text to check for monetary values.

        Returns:
            bool: True if any monetary value is found, False otherwise.
        """
        money_patterns = [
            r'\$\d{1,3}(,\d{3})*(\.\d{2})?',  # $111,111.11 or $11.1
            r'\b\d{1,3}(,\d{3})*(\.\d{2})?\s+dollars?\b',  # 11 dollars
            r'\b\d{1,3}(,\d{3})*(\.\d{2})?\s+USD\b'  # 11 USD
        ]
        pattern = '|'.join(money_patterns)
        return bool(re.search(pattern, text, re.IGNORECASE))
    
    @staticmethod
    def relative_time_to_absolute(relative_time):
        """Converts a relative time string to an absolute datetime object.

        This method takes a relative time string (e.g., '5 minutes ago', '2 hours ago')
        and converts it to an absolute datetime object.

        Args:
            relative_time (str): The relative time string to convert.

        Returns:
            datetime: The absolute datetime corresponding to the relative time.
        """
        now = datetime.now()
        match = re.match(r'·\s*(\d+)\s+(minute|minutes|hour|hours|day|days|week|weeks|month|months|year|years)\s+ago', relative_time)
        if match:
            value, unit = match.groups()
            value = int(value)
            if unit.startswith('minute'):
                return now - timedelta(minutes=value)
            elif unit.startswith('hour'):
                return now - timedelta(hours=value)
            elif unit.startswith('day'):
                return now - timedelta(days=value)
            elif unit.startswith('week'):
                return now - timedelta(weeks=value)
            elif unit.startswith('month'):
                return now - timedelta(days=value * 30)
            elif unit.startswith('year'):
                return now - timedelta(days=value * 365)
        else:
            print(f"No match for relative_time: {relative_time}")  # Debugging print
        return now  # If parsing fails, return the current time
    
    def get_element_text(self, parent, xpath):
        """Gets the text of an element if it exists, otherwise returns None."""
        try:
            return parent.find_element(By.XPATH, xpath).text
        except NoSuchElementException:
            return None
    
    def extract_article_data(self, element):
        """Extracts data from a single article element."""
        title_element = element.find_element(By.XPATH, './/h4[@class="s-title fz-16 lh-20"]/a')
        title = title_element.get_attribute('title')
        link = title_element.get_attribute('href')

        source = self.get_element_text(element, './/span[@class="s-source mr-5 cite-co"]')
        time = self.get_element_text(element, './/span[@class="fc-2nd s-time mr-8"]')
        description = self.get_element_text(element, './/p[@class="s-desc"]')
        image_url = self.get_image_url(element, title)
        
        return News(
            title=title,
            link=link,
            source=source or 'N/A',
            time=time or 'N/A',
            description=description or 'N/A',
            image_url=image_url or 'N/A'
        )
 
    def get_image_url(self, element, title):
        """Attempts to download the article's image and return its URL."""
        try:
            image_element = element.find_element(By.XPATH, './/a[@class="thmb "]/img')
            image_url = image_element.get_attribute('src')
            
            if image_url.startswith("http"):
                image_file_name = f"{title.replace(' ', '_')}_{datetime.now().strftime('%Y%m%d%H%M%S')}.jpg"
                image_save_path = os.path.join(self.pictures_dir, image_file_name)
                if self.download_image(image_url, image_save_path):
                    return os.path.basename(image_save_path)
                else:
                    self.logger.warning(f"Failed to download image from URL: {image_url}")
                    return None
            else:
                self.logger.warning(f"Invalid image URL for article: {title} - {image_url}")
                return None
        except NoSuchElementException:
            self.logger.warning(f"Image element not found for article: {title}")
            return None

    def print_articles(self):
        for index, article in enumerate(self.articles, start=1):
            print(f"Article {index}:\nTitle: {article.title}\nLink: {article.link}\nSource: {article.source}\nTime: {article.time}\nDescription: {article.description}\n")

    def save_results_to_excel(self):
        """Saves the collected articles to an Excel file.

        This method creates a new Excel workbook, adds a worksheet named 'Results',
        and populates it with the collected article data. The workbook is saved
        with a filename based on the current date and time.

        The columns in the worksheet include title, title length, whether the title
        contains money, link, source, time, description, description length, whether
        the description contains money, and image filename.

        Raises:
            Exception: If there is an error in saving the workbook.
        """
        print("Starting the creation of the workbook.")
        
        # Create a new workbook and add the search results
        excel = Files()
        excel.create_workbook(path=None)
        
        print("Workbook created successfully.")
        excel.create_worksheet('Results')
        
        # Prepare data to add to Excel
        data = [
            ['Title', 'Title Length', 'Title Contains Money', 'Link', 'Source', 'Time', 'Description', 'Description Length', 'Description Contains Money', 'Image']
        ]
        
        for article in self.articles:
            absolute_time = self.relative_time_to_absolute(article.time)
            time = absolute_time.strftime('%Y-%m-%d %H:%M:%S')
            self.logger.debug(f"Adding article to Excel: {article}")
            title_length = len(article.title)
            description_length = len(article.description)
            title_contains_money = self.contains_money(article.title)
            description_contains_money = self.contains_money(article.description)
            
            data.append([
                article.title,
                title_length,
                title_contains_money,
                article.link,
                article.source,
                time,
                article.description,
                description_length,
                description_contains_money,
                article.image_url,
            ])
        
        self.logger.debug("Data prepared for insertion into Excel.")
        excel.append_rows_to_worksheet(data, header=True, name='Results')
        self.logger.debug("Data added to worksheet 'Results'.")
        
        # Save the workbook with the current date as the filename
        file_name = f"{datetime.now().strftime('%Y-%m-%d_%H-%M-%S')}.xlsx"
        output_path = os.path.join(get_output_dir(), file_name)
        excel.save_workbook(output_path)
        self.logger.info(f"Results saved to: {output_path}")

    def wait_for_new_tab_to_load(self):
        """Waits for a new tab to load by checking the number of window handles.

        This method attempts to detect a new tab by checking the number of
        window handles every 5 seconds, up to a maximum of 6 attempts (30 seconds).

        Raises:
            AssertionError: If the new tab does not load within the expected time.
        """
        max_attempts = 6  # Check for up to 30 seconds (6 * 5s)
        for _ in range(max_attempts):
            time.sleep(5)
            if len(self.browser.get_window_handles()) > 1:
                return
        raise AssertionError("The new tab did not load within the expected time.")

    def wait_for_element_to_be_visible(self, xpath, timeout=60):
        """Waits for an element to become visible on the page.

        This method attempts to wait until the specified element is visible
        by checking every 5 seconds, up to the specified timeout period.

        Args:
            xpath (str): The XPath of the element to wait for.
            timeout (int): The maximum time to wait, in seconds (default is 60).

        Raises:
            AssertionError: If the element does not become visible within the timeout period.
        """
        max_attempts = timeout // 5  # Check every 5 seconds up to the timeout
        for _ in range(max_attempts):
            try:
                self.browser.wait_until_element_is_visible(xpath, timeout='5s')
                return
            except Exception:
                time.sleep(5)
        raise AssertionError(f"Element {xpath} not visible after {timeout} seconds.")

    def close_browser(self):
        """Closes the browser and cleans up resources."""
        try:
            self.browser.close_browser()
        except Exception as e:
            self.logger.error(f"Error closing browser: {e}")

    def switch_to_new_tab(self):
        """Switches to the newly opened browser tab.

        This method uses the switch_window method to change the context to the newly
        opened browser tab and logs the action.
        """
        self.browser.switch_window(locator='NEW')
        self.logger.info("Switched to the new tab.")
    
    def open_browser(self, url: str, word: str, months: int):
        """Opens a browser, searches for a keyword, and processes the results.

        This method opens a browser, inputs a search term, waits for the results page
        to load, switches to a new tab, finds and clicks the 'News' link, waits for the 
        news page to load, collects articles, prints them, and saves the results to an 
        Excel file.

        Args:
            url (str): The URL to open.
            word (str): The search keyword to input.
            months (int): The number of months to filter the articles by.
        """
    
        try:
            self.logger.info("Attempting to open the browser.")

            # headless
            options = create_headless_options()

            self.retry_action(lambda: self.browser.open_available_browser(url))
            self.browser.maximize_browser_window()
            self.logger.info(f"Opening URL: {url}")

            # Wait for the search button to be visible and click it
            self.logger.info("Waiting for the search box to be visible.")
            self.browser.wait_until_element_is_visible('id=ybar-sbq', timeout=180)
            
            # Locate the search box and input the search term
            search_box = self.browser.find_element('id=ybar-sbq')
            self.browser.input_text(search_box, word)
            self.browser.submit_form(search_box)
            self.logger.info(f"Input search keyword: {word}")

            # Wait until the results page is loaded in the new tab
            self.wait_for_new_tab_to_load()
            self.switch_to_new_tab()
            
            # Now look for the "News" link and click it
            self.find_and_click_news_link()

            # Wait until the news page is fully loaded
            self.wait_for_element_to_be_visible('xpath://ol[@class="mb-15 reg searchCenterMiddle"]', timeout=60)

            # Collect the articles
            self.collect_articles()

            self.filter_articles_by_date(months)

            # Print the collected articles
            self.print_articles()

            # Save the results to an Excel file
            self.save_results_to_excel()
        except AssertionError as e:
            self.logger.error(f"Error waiting for the element: {e}")
        finally:
            self.close_browser()



        
        
