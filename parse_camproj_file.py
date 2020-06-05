#!/usr/bin/python3
#
# ./parse_camproj_file.py filename
# 
# it outputs a file with list of key frames and their times
#
# G. Q. Maguire Jr.
#
# 2020.06.04
#
# commands to use:
#


import csv, requests, time
from pprint import pprint
import optparse
import sys
import os

import pathlib                  # to get each of the files

import json

from lxml import html

# Use Python Pandas to create XLSX files
import pandas as pd


def get_text_for_tag(document, tag, dir):
    tag_xpath='.//'+tag
    text_dir=tag+'_text'
    tmp_path=document.xpath(tag_xpath)
    if tmp_path:
        # remove all inline-ref and dont-index spans
        tmp=[item.text for item in tmp_path]
        tmp[:] = [item for item in tmp if item != None and item != "\n"]
        if tmp:
            dir[text_dir]=tmp

def remove_tag(document, tag):
    tag_xpath='//'+tag
    for bad in document.xpath(tag_xpath):
        bad.getparent().remove(bad)

def remove_inline_and_dont_index_tags(str1):
    document2 = html.document_fromstring(str1)
    # remove span class="inline-ref">
    for bad in document2.xpath("//span[contains(@class, 'inline-ref')]"):
        bad.getparent().remove(bad)
   # remove span class="dont-index">
    for bad in document2.xpath("//span[contains(@class, 'dont-index')]"):
        bad.getparent().remove(bad)
    return html.tostring(document2, encoding=str, method="text", pretty_print=False )

def get_editrate(document):
    global Verbose_Flag
    tmp_path=document.xpath('.//project')
    if tmp_path:
        edit_rate_string=tmp_path[0].get('editrate')
        if edit_rate_string:
            rate=edit_rate_string.split('/')[0]
            if rate:
                if Verbose_Flag:
                    print("editrate is specified as {}".format(rate))
                return int(rate)
    return 30                   # if there is no rate specified, then assume 30 frames per second

def convert_frame_number_to_seconds(f_string, rate):
    if f_string and rate > 0:
        fn=f_string.split('/')[0]
        if fn:
            return int(fn)/rate
    print("convert_frame_number_to_seconds: error no frame number or no frame rate")
    return 0

def seconds_to_time(sec):
    # convert seconds 'HH:MM:SS,FFF' - SRT uses French format for milliseconds
    m, s = divmod(sec, 60)
    h, m = divmod(m, 60)
    time_string="%02d:%02d:%06.3f" % (h, m, s)
    return time_string.replace('.', ',')

def process_project(project):
    global Verbose_Flag
    global null
    
    d=dict()

    # handle the case of an empty document
    if not project or len(project) == 0:
        return d
    
    document = html.document_fromstring(project)
    # raw_text = document.text_content()

    edit_rate=get_editrate(document)

    # process the different elements
    #
    # get the time and value each Keyframe - as this should be the title of a slide
    slide_number=1
    tmp_path=document.xpath('.//keyframe')
    if tmp_path:
        #
        for item in tmp_path:
            tmp_time=item.get('time')
            tmp_value=item.get('value')
            if tmp_time and tmp_value:
                d[slide_number]={'time': seconds_to_time(convert_frame_number_to_seconds(tmp_time, edit_rate)), 'value': tmp_value}
                slide_number=slide_number+1

    print("Last slide number={}".format(slide_number-1))
    if Verbose_Flag:
        print("project is now {}".format(html.tostring(document)))

    return d


def main():
    global Verbose_Flag
    global null

    parser = optparse.OptionParser()

    parser.add_option('-v', '--verbose',
                      dest="verbose",
                      default=False,
                      action="store_true",
                      help="Print lots of output to stdout"
    )

    options, remainder = parser.parse_args()

    Verbose_Flag=options.verbose

    if Verbose_Flag:
        print("ARGV      : {}".format(sys.argv[1:]))
        print("VERBOSE   : {}".format(options.verbose))
        print("REMAINING : {}".format(remainder))

    if (len(remainder) < 1):
        print("Inusffient arguments\n must provide name of a file to process\n")
        sys.exit()

    file_name=remainder[0]
    if Verbose_Flag:
        print("processing file {}".format(file_name))

    null = None
    project_data=dict()
        
    with open(file_name) as input_file:
        project = input_file.read()
        project_data=process_project(project)

    output_filename='keyframes-'+file_name+'.json'
    try:
        with open(output_filename, 'w') as json_file:
            json.dump(project_data, json_file)
    except:
        print("error trying to write to {}".format(output_filename))


if __name__ == "__main__": main()

