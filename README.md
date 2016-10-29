Desktop Defender II - Battle for Existence
==============================================

DD2 is a single-player 2D shoot'em up game where you control a battle turret on the surface of the moon. Your mission is to repel organized alien attacks against the Earth. The game consists of 5 separate missions, each with a set of objectives you'll need to accomplish. Players may create separate game profiles. There are statistics being collected for each profile after a mission has been completed. Stats are used for a *Hall of Fame* ranking.

# Install

## Windows

Please note that from Windows Vista onwards DirectX 7 is **no longer supported**! This game would probably only run on a Windows XP. 

*Original* system requirements:

 * Windows 98/XP
 * 300 MHz Pentium or better
 * 32 MB RAM
 * DirectX 7.0

## Linux

It should be possible to run & play DD2 on Linux using [Wine](https://www.winehq.org/).

Download the game [release package](https://github.com/kenamick/desktopdefender2/releases). Run it via wine, e.g.,

    wine ddefender2.exe
    
By default Wine will extract it to `c:\games\ddefender2`. The real path on your Linux system should be `/home/<user>/.wine/drive_c/games/ddefender2`. 

The next step is a bit tricky. The game needs DirectX 7 type descriptors for VB6, so you need to download the `dx7vb.dll` file and copy it in the game installation directory. The file may be found at [thevbzone](http://www.thevbzone.com/d_DLL.htm) or at [DllDump](http://www.dlldump.com/download-dll-files_new.php/dllfiles/D/dx7vb.dll/5.03.2600.2180/download.html).

The DLL file needs to be registered in the *Windows Registry*, so after copying it, run the following:

    wine regsvr32 dx7vb.dll

Your are now set to run the game! To start the setup run:

    wine Setup.exe
    
Choose the in-game language (English is the default) and press the `Run Game` button.

# Game Screenshots

![alt text](http://i.imgur.com/UPLz8Gr.jpg "In game #1")
![alt text](http://i.imgur.com/L005keL.jpg "In game #2")
![alt text](http://i.imgur.com/AcHIVVw.jpg "In game #3")

# Backstory

The Milky Way was always full of life. Life thrives everywhere and humans are just a minor part of Earth's history. 4 million years ago the Zonerians from Sirius, threatened by destruction, settled on our homeworld. Far from the center of the galaxy the Earth made a perfect new home for them. The Zonerians built many settlements deep underground the planet and the Moon. Following the laws of nature they decided not to interfere with Earth's evolution. They took part neither in man's wars nor in their world crises and suffering. Their socially and technologically far superior society was merely a silent observer.

The ultimate force in our galaxy is the Galactic Confederation. Most intelligent species, who managed to reach the point of space travel, are members of it. However, every great force has a great enemy. The Black League, one of the most powerful coalitions in our galaxy, is the only force capable of resisting the Confederation. After monitoring the evolution of the human species for ages, the Confederation decided to take a major step towards making Earth's human beings a part of the alliance. The Black League, however, claimed to have legal rights over the Sol and it's resources. The first step towards a galactic war was then more than clear. Securing the Sol's resources and the abundant of life planet - the Earth, was their primary goal. Having it's major forces too far away to prevent an invasion, the Galactic Confederation is helpless to stop the Black League. Only the joint human and zonerian forces stand between the Earth's total destruction.

# License

[MIT](LICENSE)
