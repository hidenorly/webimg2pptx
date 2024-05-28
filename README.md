# webimg2pptx

## requirement

```
pip install requests
pip install pillow
pip install selenium
pip install python-pptx
pip install cairosvg
pip install pyheif
#brew install libffi libheif
#pip install git+https://github.com/david-poirier-csn/pyheif.git
pip install webcolors
```

## How to use

```
% python3 webimg2pptx.py --help
usage: webimg2pptx.py [-h] [-t TEMPPATH] [-o OUTPUT] [-a] [-p] [-l LAYOUT]
                      [-f] [-w] [--minSize MINSIZE] [--maxDepth MAXDEPTH]
                      [--baseUrl BASEURL] [--timeOut TIMEOUT]
                      [--offsetX OFFSETX] [--offsetY OFFSETY]
                      [--fontFace FONTFACE] [--fontSize FONTSIZE]
                      [--title TITLE] [--titleSize TITLESIZE]
                      [--titleFormat TITLEFORMAT]
                      PAGE [PAGE ...]

Download images from web pages

positional arguments:
  PAGE                  Web pages to download images from

options:
  -h, --help            show this help message and exit
  -t TEMPPATH, --temp TEMPPATH
                        Temporary path. (default: .)
  -o OUTPUT, --output OUTPUT
                        Output PowerPoint file path (default: output.pptx)
  -a, --addUrl          Add URL to the slide (default: False)
  -p, --usePageUrl      Use page URL if possible (default: False)
  -l LAYOUT, --layout LAYOUT
                        Specify layout full or left or right (default: full)
  -f, --fullfit         Specify if want to fit within the slide (default:
                        False)
  -w, --withFullArgUrl  Specify if want to use full url with ? argument
                        (default: False)
  --minSize MINSIZE     Minimum size of images to download (format:
                        WIDTHxHEIGHT) (default: None)
  --maxDepth MAXDEPTH   maximum depth of links to follow (default: 1)
  --baseUrl BASEURL     Specify base url if you want to restrict download
                        under the baseUrl (default: )
  --timeOut TIMEOUT     Specify time out [sec] if you want to change the
                        default (default: 60)
  --offsetX OFFSETX     Specify offset x (Inch. max 16. float) (default: 0)
  --offsetY OFFSETY     Specify offset y (Inch. max 9. float) (default: 0)
  --fontFace FONTFACE   Specify font face if necessary (default: Calibri)
  --fontSize FONTSIZE   Specify font size (pt) if necessary (default: 18.0)
  --title TITLE         Specify title if necessary (default: None)
  --titleSize TITLESIZE
                        Specify title size if necessary (default: None)
  --titleFormat TITLEFORMAT
                        Specify title format if necessary e.g.
                        color:white,face:Calibri,size:40,bold (default: None)
```

```
% python3 webimg2pptx.py -t ~/tmp/test -o test.pptx --addUrl --usePageUrl --minSize=400x400 --maxDepth=2 https://hoge.com/hoge1 https://hoge.com/hoge2 --basUrl=https://hoge.com/
```