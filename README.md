# Simple-Movie-Server-1
A simple movie server for hosting and watching movies locally on a PC.

## Intro
- "Simple-Movie-Server-1" creates a webpage that allows you to use a browser to
watch movies saved to your local PC. It is super-simple to use and works great
for a Home Theater PC (HTPC) with content stored locally. Supported formats are
MP4, M4V, Mpg, Mpeg and PNG.

## Language
- VBScript
- Microsoft Windows Script Host

## Setup Instructions
1. Save your MP4, M4V, Mpg and Mpeg movies to a single folder, e.g. "movies".
2. Save PNG thumbnail images into a subdirectory called "thumbnails." Your
folder structure should appear as follows:
    ```
    movies
    ¦   movie1.mp4
    ¦   movie2.m4v
    ¦   movie3.mpg
    ¦   movie4.mpeg
    +---thumbnails
            movie1.png
            movie2.png
            movie3.png
            movie4.png
    ```
3. Save the file "Create_Movie_Page.vbs" to your computer.
4. Double-click to run the "Create_Movie_Page.vbs" script.
5. A window will popup, asking you for the location of your movies.
6. Select the folder you created and click "OK".
7. The script will create and open the new webpage.
8. Click on any link to watch the movie.

###Usage Notes
- You can move and rename the webpage and it will still work.
- You can delete the script, or run it again anytime you add or remove movies.
The webpage is always created in the same folder as the script.
- Movies and thumbnails must have the same name with different extensions.

## Troubleshooting
-I don't see any movie links. - Check that your movie titles don't contain any
unusual characters, and that your images are named the same as you movies.
- Movie doesn't work - Most likely the movie is not actually an MP4 or M4V, or
the browser does not support movie playback. Try using the most recent version
of a popular browser. If this doesn't work you may need to re-encode the video.
- "Open File - Security Warning" - This is a warning from Windows because
VBScript has the power to make changes on your PC. You should always verify and
check content downloaded from the internet beore running it on your computer.
This script is plain text and open source, so you can see for yourself that
there is no malicious code. Click "Open" to run the script.
- "Windows Script Host" error - This means my code isn't working. Drop me a
note and I'll look into a fix.
