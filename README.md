# Improvement-Mod-Launcher
Launcher for age of empires 3 improvement mod

![launcher](https://github.com/user-attachments/assets/39d7ab75-7e9f-42f7-8847-41b0ec4d6502)
Age of Empires III: Improvement Mod Launcher, Command Your Empire with Ease! <br />
Step into a world of strategic mastery with the Age of Empires III: Improvement Mod Launcher your all-in-one command center for managing the ultimate Age of Empires III Improvement Mod experience. Designed specifically for the Improvement Mod, this launcher empowers you to shape the battlefield before the first shot is fired. <br />
Seamlessly check, download, and update the mod with a single click, ensuring you’re always at the cutting edge of enhancements. Customize your strategy by switching between different AI opponents, fine-tuning the challenge to your preference. Whether you seek performance boosts or compatibility tweaks, the launcher gives you full control—toggle the DirectX 12 wrapper, enable DirectPlay, install DirectX 9, MSXML4, and essential VC++ redistributables to optimize your game. <br />
And when all is set, march into battle effortlessly launch Age of Empires III directly from the launcher and witness history unfold with the best improvements at your command. <br />
Take charge. Customize your empire. Conquer with confidence. The Age of Empires III: Improvement Mod Launcher is your gateway to the most refined AoE3 Improvement Mod experience yet!

**How to use**<br />
Install age of empires 3 (asian dynasties or gold edition), if you already have the game installed but want to keep a separate install, copy the game folder and rename it to Age of Empires III Improvement Mod.<br />
Run age3y.exe once then exit. <br />
Download the zip from [Releases](https://github.com/ageekhere/Improvement-Mod-Launcher/releases) , unzip the exe into same directory as age3y.exe  NOTE: Currently not release yet<br />
Note: The Launcher is built with python and uses autopytoexe, if your antivirus detects it as a virus it is a false positive and can be [excluded](https://nitratine.net/blog/post/issues-when-using-auto-py-to-exe/#my-antivirus-detected-the-exe-as-a-virus) <br />

**How to compile for development**<br />
Pull the git repository, you can use [GitHub Desktop](https://desktop.github.com/download) <br />
Install the latest python version (current version Python 3.13.2) and add it to your [Environment Variables](https://www.liquidweb.com/help-docs/adding-python-path-to-windows-10-or-11-path-environment-variable)<br />
Run Improvement Mod Launcher.pyw in any IDE that supports Python <br />
Install any missing modules using pip <br />

**To compile as an exe** <br /> 
pip install auto-py-to-exe <br />
Download upx for file [compression](https://github.com/upx/upx) and add it to your environment variables <br />
In cmd run autopytoexe <br />
In the autopytoexe folder from the repository find autopytoexe_config.json edit the c:/Improvement-Mod-Launcher/ locations to match your environment, note do not push your autopytoexe_config.json changes <br />
Under settings click import config from JSON file to import the config <br />
Click convert .PY to .EXE <br />
You should now have an exported .exe <br />
