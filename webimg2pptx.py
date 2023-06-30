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
import urllib.request
from urllib.parse import urljoin
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pptx import Presentation
from pptx.util import Inches, Pt

globalCache = {}


class WebPageImageDownloader:
    def getRandomFilename():
        letters = string.ascii_lowercase
        return ''.join(random.choice(letters) for i in range(10))

    def getOutputFileStream(outputPath, url):
        f = None
        filename = os.path.join(outputPath, os.path.basename(url))
        try:
            f = open(filename, 'wb')
        except:
            filename = os.path.join(outputPath, WebPageImageDownloader.getRandomFilename())
            f = open(filename, 'wb')
        filename = os.path.basename(filename)
        return f, filename

    def getImageSize(data):
        try:
            with Image.open(BytesIO(data)) as img:
                return img.size
        except:
            return None

    def downloadImage(imageUrl, outputPath, minDownloadSize=None):
        filename = None
        url = None
        if not imageUrl in globalCache:
            globalCache[imageUrl] = True

            if imageUrl.strip().endswith(".svg"):
                try:
                    with urllib.request.urlopen(imageUrl) as response:
                        svgContent = response.read()
                        url =imageUrl
                        f, filename = WebPageImageDownloader.getOutputFileStream(outputPath, imageUrl)
                        if f:
                            f.write(svgContent)
                            f.close()
                except:
                    pass
            else:
                size = None
                response = None
                try:
                    response = requests.get(imageUrl)
                    if response.status_code == 200:
                        # check image size
                        size = WebPageImageDownloader.getImageSize(response.content)
                except:
                    pass

                if response:
                    if minDownloadSize==None or (size and size[0] >= minDownloadSize[0] and size[1] >= minDownloadSize[1]):
                        url =imageUrl
                        f, filename = WebPageImageDownloader.getOutputFileStream(outputPath, imageUrl)
                        if f:
                            for chunk in response.iter_content(chunk_size=8192):
                                f.write(chunk)
                            f.close()

        return filename, url

    def isSameDomain(url1, url2, baseUrl=""):
        isSame = urlparse(url1).netloc == urlparse(url2).netloc
        isbaseUrl =  ( (baseUrl=="") or url2.startswith(baseUrl) )
        return isSame and isbaseUrl

    def _downloadImagesFromWebPage(driver, fileUrls, pageUrls, pageUrl, outputPath, minDownloadSize, baseUrl, maxDepth, depth, usePageUrl, timeOut):
        if depth > maxDepth:
            return

        element = None
        try:
            driver.get(pageUrl)
            element = WebDriverWait(driver, timeOut).until(
                EC.presence_of_element_located((By.TAG_NAME, 'a'))
            )
            element = WebDriverWait(driver, timeOut).until(
                EC.presence_of_element_located((By.TAG_NAME, 'img'))
            )
        except:
            pass

        if element:
            # download image
            for img_tag in driver.find_elements(By.TAG_NAME, 'img'):
                imageUrl = None
                try:
                    imageUrl = img_tag.get_attribute('src')
                except:
                    pass
                if imageUrl:
                    imageUrl = urljoin(pageUrl, imageUrl)
                    fileName, url = WebPageImageDownloader.downloadImage(imageUrl, outputPath, minDownloadSize)
                    if fileName and not fileName in fileUrls:
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
                        continue #print("Error occured (href is not found in a tag) at "+str(link))
                    if href and WebPageImageDownloader.isSameDomain(pageUrl, href, baseUrl):
                        if not href in pageUrls:
                            pageUrls.add(href)
                            if href.endswith(".jpg") or href.endswith(".jpeg") or href.endswith(".png") or href.endswith(".gif") or href.endswith(".svg"):
                                fileName, url = WebPageImageDownloader.downloadImage(href, outputPath, minDownloadSize)
                                if fileName and not fileName in fileUrls:
                                    if usePageUrl:
                                        fileUrls[fileName] = pageUrl
                                    elif url:
                                        fileUrls[fileName] = url
                            else:
                                WebPageImageDownloader._downloadImagesFromWebPage(driver, fileUrls, pageUrls, href, outputPath, minDownloadSize, baseUrl, maxDepth, depth + 1, usePageUrl, timeOut)


    def downloadImagesFromWebPages(urls, outputPath, minDownloadSize=None, baseUrl="", maxDepth=1, usePageUrl=False, timeOut=60):
        fileUrls = {}
        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        driver = webdriver.Chrome(options=options)
        driver.set_window_size(1920, 1080)

        pageUrls=set()
        for url in urls:
            WebPageImageDownloader._downloadImagesFromWebPage(driver, fileUrls, pageUrls, url, outputPath, minDownloadSize, baseUrl, maxDepth, 0, usePageUrl, timeOut)
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
        pic = None
        try:
            pic = self.currentSlide.shapes.add_picture(imagePath, x, y)
        except:
            pass
        if pic:
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
        return pic

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
    parser.add_argument('-t', '--temp', dest='tempPath', type=str, default='.', help='Temporary path.')
    parser.add_argument("-o", "--output", help="Output PowerPoint file path")
    parser.add_argument("-a", "--addUrl", action='store_true', default=False, help="Add URL to the slide")
    parser.add_argument("-p", "--usePageUrl", action='store_true', default=False, help="Use page URL if possible")
    parser.add_argument('--minSize', type=str, help='Minimum size of images to download (format: WIDTHxHEIGHT)')
    parser.add_argument('--maxDepth', type=int, default=1, help='maximum depth of links to follow')
    parser.add_argument('--baseUrl', type=str, default="", help='Specify base url if you want to restrict download under the baseUrl')
    parser.add_argument('--timeOut', type=int, default=60, help='Specify time out [sec] if you want to change the default')
    args = parser.parse_args()
    if args.usePageUrl:
        args.addUrl = True

    # --- download 
    minDownloadSize = None
    if args.minSize:
        minDownloadSize = tuple(map(int, args.minSize.split('x')))

    if not os.path.exists(args.tempPath):
        os.makedirs(args.tempPath)

    fileUrls = WebPageImageDownloader.downloadImagesFromWebPages(args.pages, args.tempPath, minDownloadSize, args.baseUrl, args.maxDepth, args.usePageUrl, args.timeOut)

    # --- create power point
    prs = PowerPointUtil( args.output )

    # --- sort per page url
    perPageImgFiles={}
    pageUrls = []
    for filename in fileUrls.keys():
        if filename.endswith(('.png', '.jpg', '.jpeg')):
            pageUrl = fileUrls[filename]
            if pageUrl:
                if not pageUrl in perPageImgFiles:
                    perPageImgFiles[pageUrl] = []
                    pageUrls.append(pageUrl)
                perPageImgFiles[pageUrl].append(filename)

    pageUrls.sort(key=lambda x: (len(x), x))

    # --- add image file to the slide
    for aPageUrl in pageUrls:
        for filename in perPageImgFiles[aPageUrl]:
            prs.addSlide()
            pic = prs.addPicture(os.path.join(args.tempPath, filename), 0, 0)
            if pic and args.addUrl:
                text = None
                if not args.usePageUrl:
                    text = filename
                if filename in fileUrls:
                    text = fileUrls[filename]
                if text:
                    prs.addText(text, Inches(0), Inches(PowerPointUtil.SLIDE_HEIGHT_INCH-0.4), Inches(PowerPointUtil.SLIDE_WIDTH_INCH), Inches(0.4))

    # --- save the ppt file
    prs.save()
