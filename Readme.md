#ksdust
This is a presentation remote controller dedicated to Microsoft PowerPoint on 64 bit version of Windows (at least for now)  
More features are coming in, and any kind of contribution is appreciated.


##Main Features
- read the notes of the current slide
- have a peek at the thumbnail of next slide
- go to the next/previous slide


##How to Use
- Download the bin.zip
- If your Windows cannot run any 64bit program, you'll have to download the source code and compile from scratch (See [How to Build From Source Code](#how-to-build-from-source-code))
- Open the ksdust.exe, and you must allow the app to accept any network connection.
- Open your presentation file, and switch PoewrPoint to slide show mode.
- Connect to the computer using any web browser that supports WebSocket.
- Touch/Press the upper part of the web page to go to the previous slide; the lower part of the page goes to the next slide.
- The notes of current slide will be shown on the upper part of the page, the thumbnail of the next slide will be shown on the lower part.


##How to Build From Source Code
This app is composed of three parts: 
- web server written in Go, the main function is in ```ksdust/exe/ksdust_m.go```
- PowerPoint controller library written with Visual C++ 2013 (workaround to a [weired issue](#ask-for-help))
- client web interface written in html, javascript and css.

So, to build this app from source code, you have to install
- [Go compiler](http://golang.org/doc/install) (at least 1.2)
- Visual C++ (I use version 2013)  

If you want it to run on 32 bit of Windows, be sure both dll and exe are built in 32 bit version, a mix (32 bit exe with 64 bit dll or vise versa) will FAILED to run.  

After you succeed every step of compilation process, put the PowerPointWrapper.dll, ksdust.exe and all .html, .css, .js file into the same folder.


##Known Issues
- If the presentation is opened in Protected View, the app will failed to work (For PoewrPoint 2010 and later)
- Incorrect thumbnail may shown when some slides are hidden
- Slide notes get cropped if too long
- Sometimes when you touch for next slide, you may think the app is malfunction because the notes and the thumbnail does not change, that's probably because the slide contains one or more animation. When you touch for the next slide, PowerPoint plays the animation and stay on the same slide, result in the no change of notes and thumbnail


##Future Work
- Ability to get thumbnail of EVERY step of animaion in slide, of at least a progress of animations in slide
- Redesign the client interface layout
- Get title of surrent slide
- Make a interface for iPad (portrait)
- Refactor the code


##Ask For Help
- At first, I communicate with PowerPoint using the [go-ole](https://github.com/mattn/go-ole/) COM helper library. It worked quite well until I want to get an element from array using the equivalent syntax below:  

```vb
Dim pptApp As PowerPoint.Application  
Dim slide = pptApp.ActivePresentation.Slides(1) 'Get first slide  
```

I always get a "Member not found" error with it, which drove my crazy. Finally I had to do the commuication using a external library call (with the help of [DispHelper](http://zh.sourceforge.jp/projects/sfnet_disphelper/)).  
If someone can help with that, I will be able to turn this app into a pure Go implementation without the help of a external dll. Again any help will be appreciated.


#Something Written in Chinese
這是我學習 Go 兩個多月來第一次寫 web app，希望這個作品可以引領我進入更深的 Go 世界。
