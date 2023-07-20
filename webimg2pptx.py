#   Copyright 2023 hidenorly
#
#   Licensed under the Apache License, Version 2.0 (the "License");
#   you may not use this file except in compliance with the License.
#   You may obtain a copy of the License at
#
#       http://www.apache.org/licenses/LICENSE-2.0
#
#   Unless required by applicable law or agreed to in writing, software
#   distributed under the License is distributed on an "AS IS" BASIS,
#   WITHOUT WARRANTIES OR CONDITIONS OF ANY KIND, either express or implied.
#   See the License for the specific language governing permissions and
#   limitations under the License.

import argparse
import os
import re
import random
import requests
import string
import random
import time

from PIL import Image
from io import BytesIO
import cairosvg
import pyheif
import urllib.request
from urllib.parse import urljoin
from urllib.parse import urlparse
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.support.ui import WebDriverWait
from selenium.webdriver.support import expected_conditions as EC

from pptx import Presentation
from pptx.util import Inches, Pt
from pptx.enum.text import PP_ALIGN
from pptx.enum.text import MSO_ANCHOR

globalCache = {}

class ImageUtil:
    def getFilenameWithExt(filename, ext=".jpeg"):
        filename = os.path.splitext(filename)[0]
        return filename + ext

    def covertToJpeg(imageFile):
        outFilename = ImageUtil.getFilenameWithExt(imageFile, ".jpeg")
        image = None
        if imageFile.endswith(('.heic', '.HEIC')):
            try:
                heifImage = pyheif.read(imageFile)
                image = Image.frombytes(
                    heifImage.mode,
                    heifImage.size,
                    heifImage.data,
                    "raw",
                    heifImage.mode,
                    heifImage.stride,
                )
            except:
                pass
        else:
            try:
                image = Image.open(imageFile)
            except:
                pass
        if image:
            image.save(outFilename, "JPEG")
        return outFilename

    def getImageSize(imageFile):
        try:
            with Image.open(imageFile) as img:
                return img.size
        except:
            return None

    def getImageSizeFromChunk(data):
        try:
            with Image.open(BytesIO(data)) as img:
                return img.size
        except:
            return None

    def convertSvgToPng(svgPath, pngPath, width=1920, height=1080):
        try:
            cairosvg.svg2png(url=svgPath, write_to=pngPath, output_width=width, output_height=height)
        except:
            pass



class WebPageImageDownloader:
    def getRandomFilename():
        letters = string.ascii_lowercase
        return ''.join(random.choice(letters) for i in range(10))

    def getSanitizedFilenameFromUrl(url):
        parsed_url = urllib.parse.urlparse(url)
        filename = parsed_url.path.split('/')[-1]

        filename = re.sub(r'[\\/:*?"<>|]', '', filename)

        return filename

    def getOutputFileStream(outputPath, url):
        f = None
        filename = WebPageImageDownloader.getSanitizedFilenameFromUrl(url)
        filename = str(os.path.join(outputPath, filename))
        if not filename.endswith(('.png', '.jpg', '.jpeg', '.svg', '.gif')):
            filename = filename+".jpeg"

        try:
            f = open(filename, 'wb')
        except:
            filename = os.path.join(outputPath, WebPageImageDownloader.getRandomFilename())
            f = open(filename, 'wb')
        filePath = filename
        filename = os.path.basename(filename)
        return f, filename, filePath


    def downloadImage(imageUrl, outputPath, minDownloadSize=None):
        filename = None
        url = None
        if not imageUrl in globalCache:
            globalCache[imageUrl] = True
            filePath = None

            if imageUrl.strip().endswith((".heic", ".HEIC", ".svg")):
                try:
                    with urllib.request.urlopen(imageUrl) as response:
                        imgContent = response.read()
                        url =imageUrl
                        f, filename, filePath = WebPageImageDownloader.getOutputFileStream(outputPath, imageUrl)
                        if f:
                            f.write(imgContent)
                            f.close()
                except:
                    pass

                if os.path.exists(filePath):
                    if imageUrl.strip().endswith((".svg")):
                        newPngPath = filePath+".png"
                        ImageUtil.convertSvgToPng(filePath, newPngPath)
                        if os.path.exists(newPngPath):
                            filename = newPngPath
                    else:
                        # .heic, .HEIC
                        newJpegPath = ImageUtil.covertToJpeg(filePath)
                        if os.path.exists(newJpegPath):
                            size = ImageUtil.getImageSize(newJpegPath)
                            if minDownloadSize==None or (size and size[0] >= minDownloadSize[0] and size[1] >= minDownloadSize[1]):
                                filename = newJpegPath
            else:
                # .png, .jpeg, etc.
                size = None
                response = None
                try:
                    response = requests.get(imageUrl)
                    if response.status_code == 200:
                        # check image size
                        size = ImageUtil.getImageSizeFromChunk(response.content)
                except:
                    pass

                if response:
                    if minDownloadSize==None or (size and size[0] >= minDownloadSize[0] and size[1] >= minDownloadSize[1]):
                        url =imageUrl
                        f, filename, filePath = WebPageImageDownloader.getOutputFileStream(outputPath, imageUrl)
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
            if not pageUrl in globalCache:
                #globalCache[pageUrl] = True
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
                            if href.endswith(('.png', '.jpg', '.jpeg', '.svg', '.gif')):
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
        tempDriver = webdriver.Chrome(options=options)
        userAgent = tempDriver.execute_script("return navigator.userAgent")
        userAgent = userAgent.replace("headless", "")
        userAgent = userAgent.replace("Headless", "")

        options = webdriver.ChromeOptions()
        options.add_argument('--headless')
        options.add_argument(f"user-agent={userAgent}")
        driver = webdriver.Chrome(options=options)
        driver.set_window_size(1920, 1080)

        pageUrls=set()
        for url in urls:
            WebPageImageDownloader._downloadImagesFromWebPage(driver, fileUrls, pageUrls, url, outputPath, minDownloadSize, baseUrl, maxDepth, 0, usePageUrl, timeOut)
        driver.quit()
        tempDriver.quit()

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

    # layout is full, left, right, top, bottom
    def getLayoutPosition(self, layout="full"):
        # for full
        x=0
        y=0
        width = self.prs.slide_width
        height = self.prs.slide_height

        if layout=="left" or layout=="right":
            width = width /2
        if layout=="top" or layout=="bottom":
            height = height /2
        if layout=="right":
            x=width
        if layout=="bottom":
            y=height

        return x,y,width,height

    def getLayoutToFitRegion(self, width, height, regionWidth, regionHeight):
        resultWidth = width
        resultHeight = height

        if width > height:
            resultWidth = regionWidth
            resultHeight = int(regionWidth * height / width+0.99)
        else:
            resultHeight = regionHeight
            resultWidth = int(regionHeight * width / height+0.99)

        return resultWidth, regionHeight


    def addSlide(self, layout=None):
        if layout == None:
            layout = self.prs.slide_layouts[6]
        self.currentSlide = self.prs.slides.add_slide(layout)

    def addPicture(self, imagePath, x=0, y=0, width=None, height=None, isFitToSlide=True, regionWidth=None, regionHeight=None, isFitWihthinRegion=False):
        if not regionWidth:
            regionWidth = self.prs.slide_width
        if not regionHeight:
            regionHeight = self.prs.slide_height
        regionWidth = int(regionWidth+0.99)
        regionHeight = int(regionHeight+0.99)
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
                    picWidth = pic.width
                    picHeight = pic.height
                    if width > height:
                        picWidth = regionWidth
                        picHeight = int(regionWidth * height / width + 0.99)
                    else:
                        picHeight = regionHeight
                        picWidth = int(regionHeight * width / height + 0.99)
                    if isFitWihthinRegion:
                        deltaWidth = picWidth - regionWidth
                        deltaHeight = picHeight - regionHeight
                        if deltaWidth>0 or deltaHeight>0:
                            # exceed the region
                            if deltaWidth > deltaHeight:
                                picWidth = regionWidth
                                picHeight = int(regionWidth * height / width + 0.99)
                            else:
                                picHeight = regionHeight
                                picWidth = int(regionHeight * width / height + 0.99)
                    pic.width = picWidth
                    pic.height = picHeight
        return pic

    def addText(self, text, x=Inches(0), y=Inches(0), width=None, height=None, fontFace='Calibri', fontSize=Pt(18), isAdjustSize=True, textAlign = PP_ALIGN.LEFT, isVerticalCenter=False):
        if width==None:
            width=self.prs.slide_width
        if height==None:
            height=self.prs.slide_height
        width = int(width+0.99)
        height = int(height+0.99)

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

        if isVerticalCenter:
            text_frame.vertical_anchor = MSO_ANCHOR.MIDDLE

        for paragraph in text_frame.paragraphs:
            paragraph.alignment = textAlign



if __name__ == '__main__':
    parser = argparse.ArgumentParser(description='Download images from web pages', formatter_class=argparse.ArgumentDefaultsHelpFormatter)
    parser.add_argument('pages', metavar='PAGE', type=str, nargs='+', help='Web pages to download images from')
    parser.add_argument('-t', '--temp', dest='tempPath', type=str, default='.', help='Temporary path.')
    parser.add_argument("-o", "--output", help="Output PowerPoint file path")
    parser.add_argument("-a", "--addUrl", action='store_true', default=False, help="Add URL to the slide")
    parser.add_argument("-p", "--usePageUrl", action='store_true', default=False, help="Use page URL if possible")
    parser.add_argument("-l", "--layout", action='store', default="full", help="Specify layout full or left or right")
    parser.add_argument("-f", "--fullfit", action='store_true', default=False, help="Specify if want to fit within the slide")
    parser.add_argument('--minSize', type=str, help='Minimum size of images to download (format: WIDTHxHEIGHT)')
    parser.add_argument('--maxDepth', type=int, default=1, help='maximum depth of links to follow')
    parser.add_argument('--baseUrl', type=str, default="", help='Specify base url if you want to restrict download under the baseUrl')
    parser.add_argument('--timeOut', type=int, default=60, help='Specify time out [sec] if you want to change the default')
    parser.add_argument('--offsetX', type=float, default=0, help='Specify offset x (Inch. max 16. float)')
    parser.add_argument('--offsetY', type=float, default=0, help='Specify offset y (Inch. max 9. float)')
    parser.add_argument('--fontFace', type=str, default="Calibri", help='Specify font face if necessary')
    parser.add_argument('--fontSize', type=float, default=18.0, help='Specify font size (pt) if necessary')
    parser.add_argument('--title', type=str, default=None, help='Specify title if necessary')
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
        if filename.endswith(('.png', '.jpg', '.jpeg', '.svg', '.gif')):
            pageUrl = fileUrls[filename]
            if pageUrl:
                if not pageUrl in perPageImgFiles:
                    perPageImgFiles[pageUrl] = []
                    pageUrls.append(pageUrl)
                perPageImgFiles[pageUrl].append(filename)

    pageUrls.sort(key=lambda x: (len(x), x))

    # --- add image file to the slide
    x, y, regionWidth, regionHeight = prs.getLayoutPosition(args.layout)
    offsetX = Inches(args.offsetX)
    offsetY = Inches(args.offsetY)
    x = x + offsetX
    y = y + offsetY
    regionWidth = int( regionWidth - offsetX )
    regionHeight = int( regionHeight - offsetY )
    isFitWihthinRegion = args.fullfit
    fontFace = args.fontFace
    fontSize = Pt(args.fontSize)
    textAlign = PP_ALIGN.LEFT
    if args.layout == "right":
        textAlign = PP_ALIGN.RIGHT
    titleSize = args.offsetY*72.0
    if titleSize<100 or titleSize>400000:
        titleSize = Pt(40)

    for aPageUrl in pageUrls:
        for filename in perPageImgFiles[aPageUrl]:
            prs.addSlide()
            pic = prs.addPicture(os.path.join(args.tempPath, filename), x, y, None, None, True, regionWidth, regionHeight, isFitWihthinRegion)
            # Add Title
            if args.title:
                prs.addText(args.title, x, 0, regionWidth, offsetY, fontFace, titleSize, True, textAlign, True)
            # Add filename(URL) at bottom
            if pic and args.addUrl:
                text = None
                if not args.usePageUrl:
                    text = filename
                if filename in fileUrls:
                    text = fileUrls[filename]
                if text:
                    prs.addText(text, x, int(y+regionHeight-Inches(0.4)), regionWidth, Inches(0.4), fontFace, fontSize, True, textAlign)

    # --- save the ppt file
    prs.save()
