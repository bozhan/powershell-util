# powershell-util
Various powershell utility functions and scripts used in other poweshell project.

#Get-MediaInfo
Reports a summary of the video files in each folder found in the provided path.
Recurse switch -r could be used here if you want to report on subfolders.

The summary consists of several attributes:
MB/Min - ratio of size in MB to lengt of the video (helful if you want to figure out if a media file can be compressed further with e.g. h264/h265 without significant quality loss)
Size(MB) - size of the folder containing media files (only media files are considered in this calcualtion) 
Length(Min) - length of the video in minutes
#Files - number of files in each reported folder
Video BR - average video bitrate for all files in a listed folder
Audio BR - average audio bitrate for all files in a listed folder
Frame Width - the average video fram width for all files in a listed folder
Forlder Name - the folder name listed
Forlder Parent - the parent folder of the listed folder name

For more information on the syntax and options avaialbe use the -help switch on the cmdlet.