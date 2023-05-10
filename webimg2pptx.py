#   Copyright 2023 hidenorly
#
#   Licensed baseUrl the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed baseUrl the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations baseUrl the License.

import argparse
import os
import random
import requests
import string
import random
import time

from PIL import Image
from io import BytesIO
from urllib.parse import urljoin
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pptx import Presentation
from pptx.util import Inches, Pt


class WebPageImageDownloader:
    def getRandomFilename():
        letters = string.ascii_lowercase
        return ''.join(random.choice(letters) for i in range(10))

    def getImageSize(data):
        try:
            with Image.open(BytesIO(data)) as img:
                return img.size
        except Exception:
            return None

    def downloadImage(imageUrl, outputPath, minDownloadSize=None):
        filename = None
        url = None
        response = requests.get(imageUrl)
        if response.status_code == 200:
            # check image size
            size = WebPageImageDownloader.getImageSize(response.content)

            if minDownloadSize==None or (size and size[0] >= minDownloadSize[0] and size[1] >= minDownloadSize[1]):
                url =imageUrl
                filename = os.path.join(outputPath, os.path.basename(imageUrl))
                f = None
                try:
                    f = open(filename, 'wb')
                except:
                    filename = os.path.join(outputPath, WebPageImageDownloader.getRandomFilename())
                    f = open(filename, 'wb')
                if f:
                    for chunk in response.iter_content(chunk_size=8192):
                        f.write(chunk)
        if filename:
            filename = os.path.basename(filename)
        return filename, url

    def isSameDomain(url1, url2, baseUrl=""):
        isSame = urlparse(url1).netloc == urlparse(url2).netloc
        isbaseUrl =  ( (baseUrl=="") or url2.startswith(baseUrl) )
        return isSame and isbaseUrl

    def downloadImagesFromWebPage_(driver, fileUrls, pagesUrls, pageUrl, outputPath, minDownloadSize=None, baseUrl="", maxDepth=1, depth=0, usePageUrl=False):
        if depth > maxDepth:
            return

        driver.get(pageUrl)
        element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, 'a'))
        )
        element = WebDriverWait(driver, 30).until(
            EC.presence_of_element_located((By.TAG_NAME, 'img'))
        )
        # download image
        for img_tag in driver.find_elements(By.TAG_NAME, 'img'):
            imageUrl = img_tag.get_attribute('src')
            if imageUrl:
                imageUrl = urljoin(pageUrl, imageUrl)
                fileName, url = WebPageImageDownloader.downloadImage(imageUrl, outputPath, minDownloadSize)
                if fileName:
                    if usePageUrl:
                        fileUrls[fileName] = pageUrl
                    elif url:
                        fileUrls[fileName] = url

        # get links to other pages
        links = driver.find_elements(By.TAG_NAME, 'a')
        for link in links:
            if link:
                href = None
                try:
                    href = link.get_attribute('href')
                except:
                    print("Error occured (href is not found in a tag) at "+str(link))
                if href and WebPageImageDownloader.isSameDomain(pageUrl, href, baseUrl):
                    oldLen= len(pagesUrls)
                    pagesUrls.add(href)
                    if len(pagesUrls)>oldLen:
                        if href.endswith(".jpg") or href.endswith(".jpeg") or href.endswith(".png"):
                            fileName, url = WebPageImageDownloader.downloadImage(href, outputPath, minDownloadSize)
                            if fileName:
                                if usePageUrl:
                                    fileUrls[fileName] = pageUrl
                                elif url:
                                    fileUrls[fileName] = url
                        else:
                            WebPageImageDownloader.downloadImagesFromWebPage_(driver, fileUrls, pagesUrls, href, outputPath, minDownloadSize, baseUrl, maxDepth, depth + 1, usePageUrl)

    def downloadImagesFromWebPage(url, outputPath, minDownloadSize=None, baseUrl="", maxDepth=1, usePageUrl=False):
        fileUrls = {}
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)
        driver.set_window_size(1920, 1080)

        pagesUrls=set()
        WebPageImageDownloader.downloadImagesFromWebPage_(driver, fileUrls, pagesUrls, url, outputPath, minDownloadSize, baseUrl, maxDepth, 0, usePageUrl)
        driver.quit()
        return fileUrls


class PowerPointUtil:
    SLIDE_WIDTH_INCH = 16
    SLIDE_HEIGHT_INCH = 9

    def __init__(self, path):
        self.prs = Presentation()
        self.prs.slide_width  = Inches(self.SLIDE_WIDTH_INCH)
        self.prs.slide_height = Inches(self.SLIDE_HEIGHT_INCH)
        self.path = path

    def save(self):
        self.prs.save(self.path)

    def addSlide(self, layout=None):
        if layout == None:
            layout = self.prs.slide_layouts[6]
        self.currentSlide = self.prs.slides.add_slide(layout)

    def addPicture(self, imagePath, x=0, y=0, width=None, height=None, isFitToSlide=True):
        pic = self.currentSlide.shapes.add_picture(imagePath, x, y)
        if width and height:
            pic.width = width
            pic.height = height
        else:
            if isFitToSlide:
                width, height = pic.image.size
                if width > height:
                    pic.width = Inches(self.SLIDE_WIDTH_INCH)
                    pic.height = Inches(self.SLIDE_WIDTH_INCH * height / width)
                else:
                    pic.height = Inches(self.SLIDE_HEIGHT_INCH)
                    pic.width = Inches(self.SLIDE_HEIGHT_INCH * width / height)

    def addText(self, text, x=Inches(0), y=Inches(0), width=None, height=None, fontFace='Calibri', fontSize=Pt(18), isAdjustSize=True):
        if width==None:
            width=Inches(self.SLIDE_WIDTH_INCH)
        if height==None:
            height=Inches(self.SLIDE_HEIGHT_INCH)

        textbox = self.currentSlide.shapes.add_textbox(x, y, width, height)
        text_frame = textbox.text_frame
        text_frame.text = text
        font = text_frame.paragraphs[0].font
        font.name = fontFace
        font.size = fontSize
        theHeight = textbox.height
        
        if isAdjustSize:
            text_frame.auto_size = True
            textbox.top = y



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Download images from web pages')
    parser.add_argument('pages', metavar='PAGE', type=str, nargs='+', help='Web pages to download images from')
    parser.add_argument('-t', '--temp', dest='tempPath', type=str, default='.', help='Temporary path')
    parser.add_argument("-o", "--output", help="Output PowerPoint file path")
    parser.add_argument("-a", "--addUrl", action='store_true', default=False, help="Add URL to the slide")
    parser.add_argument("-p", "--usePageUrl", action='store_true', default=False, help="Use page URL if possible")
    parser.add_argument('--minSize', type=str, help='Minimum size of images to download (format: WIDTHxHEIGHT)')
    parser.add_argument('--maxDepth', type=int, default=1, help='maximum depth of links to follow')
    parser.add_argument('--baseUrl', type=str, default="", help='Specify base url if you want to restrict download under the baseUrl')
    args = parser.parse_args()
    if args.usePageUrl:
        args.addUrl = True


    # --- download
    minDownloadSize = None
    fileUrls={}
    if args.minSize:
        minDownloadSize = tuple(map(int, args.minSize.split('x')))

    if not os.path.exists(args.tempPath):
        os.makedirs(args.tempPath)

    for page in args.pages:
        fileUrls = WebPageImageDownloader.downloadImagesFromWebPage(page, args.tempPath, minDownloadSize, args.baseUrl, args.maxDepth, args.usePageUrl)

    # --- create power point
    prs = PowerPointUtil( args.output )

    for dirpath, dirnames, filenames in os.walk(args.tempPath):
        for filename in filenames:
            if filename.endswith(('.png', '.jpg', '.jpeg')):
                prs.addSlide()
                prs.addPicture(os.path.join(dirpath, filename), 0, 0)
                if args.addUrl:
                    text = filename
                    if filename in fileUrls:
                        text = fileUrls[filename]
                    prs.addText(text, Inches(0), Inches(PowerPointUtil.SLIDE_HEIGHT_INCH-0.4), Inches(PowerPointUtil.SLIDE_WIDTH_INCH), Inches(0.4))

    prs.save()
