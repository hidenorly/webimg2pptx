#   Copyright 2023, 2024 hidenorly
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

import os

from PIL import Image
from io import BytesIO
import cairosvg
import pyheif
try:
    import pillow_avif
except:
    pass

import webcolors

class ImageUtil:
    def getFilenameWithExt(filename, ext=".jpeg"):
        filename = os.path.splitext(filename)[0]
        return filename + ext

    def getImage(imageFile):
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
        return image

    def covertToJpeg(imageFile):
        outFilename = ImageUtil.getFilenameWithExt(imageFile, ".jpeg")
        image = ImageUtil.getImage(imageFile)
        if image:
            try:
                image = image.convert('RGB')
                image.save(outFilename, "JPEG")
            except:
                pass
        return outFilename

    def covertToPng(imageFile):
        outFilename = ImageUtil.getFilenameWithExt(imageFile, ".png")
        image = ImageUtil.getImage(imageFile)
        if image:
            try:
                image.save(outFilename, "PNG")
            except:
                pass
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
