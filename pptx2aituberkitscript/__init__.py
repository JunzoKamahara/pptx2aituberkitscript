import collections
import collections.abc
import os
import sys
from typing import Union
from pptx2aituberkitscript.tools import copy_file

from pptx import Presentation

import pptx2aituberkitscript.outputter as outputter
from pptx2aituberkitscript.global_var import g
from pptx2aituberkitscript.parser import parse


# initialization functions
def __prepare_titles(title_path):
    with open(title_path, "r", encoding="utf8") as f:
        indent = -1
        for line in f.readlines():
            cnt = 0
            while line[cnt] == " ":
                cnt += 1
            if cnt == 0:
                g.titles[line.strip()] = 1
            else:
                if indent == -1:
                    indent = cnt
                    g.titles[line.strip()] = 2
                else:
                    g.titles[line.strip()] = cnt // indent + 1
                    g.max_custom_title = max([g.max_custom_title, cnt // indent + 1])


def convert(
    pptx_path: str,  # path to the pptx file to be converted
    title: Union[str, None] = None,  # path to the custom title list file
    output_dir: Union[str, None] = None,  # path of the output file
    image_dir: Union[str, None] = None,  # where to put images extracted
    image_width: Union[int, None] = None,  # maximum image with in px
    disable_image: bool = False,  #  disable image extraction
    disable_wmf: bool = False,  # keep wmf formatted image untouched(avoid exceptions under linux)
    disable_color: bool = False,  # do not add color HTML tags
    disable_escaping: bool = False,  # do not attempt to escape special characters
    enable_slides: bool = False,  # add slide deliniation
    min_block_size: int = 15,  # the minimum character number of a text block to be converted
    page: Union[int, None] = None # only convert the specified page 
):

    file_path = pptx_path
    g.file_prefix = "".join(os.path.basename(file_path).split(".")[:-1])
    g.output_dir = os.path.join(os.path.dirname(file_path), g.file_prefix)

    if title == str:
        g.use_custom_title
        __prepare_titles(title)
        g.use_custom_title = True

    out_path = os.path.join(g.output_dir,"slides.md")
    script_path = os.path.join(g.output_dir,'scripts.json')

    if output_dir:
        g.output_dir = output_dir
        out_path = os.path.join(g.output_dir,out_path)
        script_path = os.path.join(g.output_dir,script_path)

    if os.path.exists(g.output_dir)==False:
        os.makedirs(g.output_dir)

    g.out_path = os.path.abspath(out_path)
    g.script_path = os.path.abspath(script_path)
    g.img_path = os.path.join(g.output_dir, "images")

    if image_dir:
        g.img_path = image_dir

    g.img_path = os.path.abspath(g.img_path)

    if image_width:
        g.max_img_width = image_width

    if min_block_size:
        g.text_block_threshold = min_block_size

    if disable_image:
        g.disable_image = True
    else:
        g.disable_image = False

    if disable_wmf:
        g.disable_wmf = True
    else:
        g.disable_wmf = False

    if disable_color:
        g.disable_color = True
    else:
        g.disable_color = False

    if disable_escaping:
        g.disable_escaping = True
    else:
        g.disable_escaping = False

    g.enable_slides = True # default to True
    if page is not None:
        g.page = page

    if not os.path.exists(file_path):
        print(f"source file {file_path} not exist!")
        print(f"absolute path: {os.path.abspath(file_path)}")
        sys.exit(1)
    prs = Presentation(file_path)
    out = outputter.md_outputter(out_path, script_path)
    parse(prs, out)
    copy_file(os.path.join(os.path.abspath(__file__),'thema.css'), 
                  os.path.join(g.output_dir, 'thema.css'))
