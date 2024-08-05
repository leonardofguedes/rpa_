import time
from robocorp.tasks import task
from custom import CustomSelenium
from robocorp import workitems

@task
def minimal_task():
    max_retries = 5  # Number of retries in case of failure
    start_time = time.time()

    item = workitems.inputs
    payload = item.current.payload if item.current else None
    
    search_phrase = payload.get('search_phrase', 'car sale increase') if payload else 'car sale increase'
    months = payload.get('months', 1) if payload else 1
    
    for attempt in range(max_retries):
        try:
            selenium = CustomSelenium()
            selenium.open_browser('https://news.yahoo.com/', search_phrase, months)
            # print("Done.")
            break  
        except Exception as e:
            print(f"Attempt {attempt + 1} failed: {e}")
            if attempt == max_retries - 1:
                print("Max retries reached. Task failed.")
    end_time = time.time()
    total_time = end_time - start_time
    print(f"Task completed in {attempt + 1} attempts and {total_time:.2f} seconds.")

if __name__ == "__main__":
    minimal_task()
