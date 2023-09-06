# 셀레니움 import
from bs4 import BeautifulSoup
from selenium import webdriver
from selenium.webdriver.common.by import By
import time
# 빈  chrome 실행
driver = webdriver.Chrome()
# 구글 사이트로 이동
# driver.get("https://www.google.com")
driver.get("https://www.naver.com")
# 뒤로 가기
driver.back()

# 앞으로 가기
driver.forward()

# 웹 페이지에서 ID 속성을 사용하여 요소를 찾기
element = driver.find_element(By.ID, 'my-id')
# 웹 페이지에서 NAME 속성을 사용하여 요소를 찾기
element = driver.find_element(By.NAME, 'my-name')
# 웹 페이지에서 CLASS_NAME 속성을 사용하여 요소를 찾기
element = driver.find_element(By.CLASS_NAME, 'my-class')
# 웹 페이지에서 TAG_NAME 속성을 사용하여 요소를 찾기
element = driver.find_element(By.TAG_NAME, 'div')
# 웹 페이지에서 LINK_TEXT 속성을 사용하여 요소를 찾기
element = driver.find_element(By.LINK_TEXT, 'my-link-text')
# 웹 페이지에서 링크 텍스트의 부분 문자열을 사용하여 사용하여 요소를 찾기
element = driver.find_element(By.PARTIAL_LINK_TEXT, 'my-link')
# 웹 페이지에서 XPATH 사용하여 요소를 찾기
element = driver.find_element(By.XPATH, '//div[@class="my-class"]')


#요소 클릭 하기 
element.click()

# 스크롤 다운
driver.execute_script("window.scrollTo(0, document.body.scrollHeight);")
time.sleep(3)

# 스크롤 업
driver.execute_script("window.scrollTo(0, 0);")