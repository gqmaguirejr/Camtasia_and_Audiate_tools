#!/usr/bin/python3
#
# ./srt2spreadsheet.py filename
# 
# it outputs a spreadsheet with the captions from the SRT file
#
# G. Q. Maguire Jr.
#
# 2020.06.05
#
# commands to use:
# ./srt2spreadsheet.py filename
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

    null = None
    srt_data=dict()
        
    state='find_caption_number'
    caption_text=''

    with open(file_name) as input_file:
        for cnt, line in enumerate(input_file):
            if Verbose_Flag:
                print("state: {0}, line length: {1} - Line {2}: {3}".format(state, len(line), cnt, line))
            if len(line) == 1 and line == '\n':
                srt_data[caption_number] = {'frame': caption_timestamp*rate,
                                            'time': caption_timestamp,
                                            'duration': caption_duration,
                                            'text': caption_text.rstrip() 
                }
                state='find_caption_number'
                caption_text=''
                continue

            if state == 'find_caption_number':
                # read the caption number
                caption_number=int(line)
                state='get_time_info'
                continue

            if state == 'get_time_info':
                arrow=line.find('-->') 
                if arrow > 0:
                    # found a time stamp line
                    caption_timestamp_string=line[0:arrow-1].replace(',', '.')
                    caption_duration_string=line[arrow+3:].replace(',', '.')
                    caption_timestamp=timestring_to_seconds(caption_timestamp_string)
                    caption_duration=timestring_to_seconds(caption_duration_string)
                    state='get_caption_text'
                else:
                    print("Error: unable to find expected timestamp")
                    state='find_caption_number'
                continue
                    
            if state == 'get_caption_text':
                if len(caption_text) == 0:
                    caption_text=line
                else:
                    caption_text=caption_text+'\n'+line
                continue


    print("caption_number={0}, time={1}, duration={2}, text={3}".format(caption_number, caption_timestamp, caption_duration, caption_text))

    # # store the last caption
    # srt_data[caption_number] = {'frame': caption_timestamp*rate,
    #                             'time': caption_timestamp,
    #                             'duration': caption_duration,
    #                             'text': caption_text.rstrip() 
    # }


    if Verbose_Flag:
        print("srt_data={0}".format(srt_data))

    captions_df=pd.DataFrame(srt_data)

    # set up the output write
    writer = pd.ExcelWriter('captions-'+file_name+'.xlsx', engine='xlsxwriter')

    captions_df.T.to_excel(writer, sheet_name='Captions')

    # Close the Pandas Excel writer and output the Excel file.
    writer.save()



if __name__ == "__main__": main()

