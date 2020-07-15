#!/usr/bin/python3
#
# ./spreadsheet2srt.py filename
# 
# the input is a spreadsheet with the captions and the output is a SRT file
#
# G. Q. Maguire Jr.
#
# 2020.06.05
#
# commands to use:
# ./spreadsheet2srt.py filename filename
#
# The spreadsheet is assumed to have the following columns: 'frame', 'time', 'duration', 'text'
#
# Note that in the SRT to spreadsheet program we put the caption number in the first column ("A"),
# but this value is ignored when reading the spreadsheet so rows can simply be deleted.
# This will mean that the caption numbering will be changed - so deleting rows will produce a file that does not match the original SRT file.
#
#
# The entries in the SRT file have the form:
#1
# 00:00:00,000 --> 00:00:02,240
# welcome to this module on ethical research
# 2
#...
#
# The list with a number is the Nth caption
# the first time is the time in the form_ HH:MM:SS,xxx where xxx is the milliseconds
#  Note that a "," is used for the decimal place; as the format was originally developed in France
# the second time is the duration of the clip (i.e., until when this caption is to be shown)
# The next line is the caption itself. The caption may be multiple lines and ends with a blank line
#


import csv, requests, time
from pprint import pprint
import optparse
import sys
import os

import json

# Use Python Pandas to create XLSX files
import pandas as pd

import math                     # needed for isnan() test

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

def timestring_to_seconds(time_string):
    # convert 'HH:MM:SS.FFF' to seconds
    split_time=time_string.split(':')
    h=int(split_time[0])
    m=int(split_time[1])
    s=float(split_time[2])
    return (h*3600 + m*60 + s)
    


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

    parser.add_option('-r', '--rate',
                      dest="rate",
                      default=30,
                      type="int",
                      help="Print lots of output to stdout"
    )


    options, remainder = parser.parse_args()

    Verbose_Flag=options.verbose
    rate=options.rate      # a rate in frames per second

    if Verbose_Flag:
        print("ARGV      : {}".format(sys.argv[1:]))
        print("VERBOSE   : {}".format(options.verbose))
        print("REMAINING : {}".format(remainder))

    if (len(remainder) < 1):
        print("Inusffient arguments\n must provide name of a file to process\n")
        sys.exit()

    file_name=remainder[0]
    if Verbose_Flag:
        print("processing file {0}".format(file_name))

    # read the sheet of Students in
    captions_df = pd.read_excel(open(file_name, 'rb'), sheet_name='Captions')
       
    output_file_name=file_name+'.srt'

    with open(output_file_name, 'w') as output_file:
        # The columns of the dataframe are 'frame', 'time', 'duration', 'text'
        for index, row in captions_df.iterrows():
            if Verbose_Flag:
                print("index: {0} 'row: {1},{2},{3},{4}".format(index, row['frame'], row['time'], row['duration'], row['text']))
            if math.isnan(row['frame']):
                break
            caption_number=index+1
            print("{0}".format(caption_number), file=output_file)

            caption_timestamp_string=seconds_to_time(row['time'])
            caption_duration_string=seconds_to_time(row['duration'])
            print("{0} --> {1}".format(caption_timestamp_string, caption_duration_string), file=output_file)
                    
            caption_text=row['text']
            print("{0}\n".format(caption_text), file=output_file)


if __name__ == "__main__": main()

