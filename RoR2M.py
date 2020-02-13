import win32com.client
import subprocess
import threading
import requests
import shutil
import time
import json
import uuid
import vdf
import sys
import os

from zipfile import ZipFile
from PyQt5.QtWidgets import QApplication, QWidget, QPushButton, QVBoxLayout, QHBoxLayout,\
    QLineEdit, QFileDialog

class IOManager: ## Manages reading and writing data to files.
    def __init__(self, file, start=True, jtype=True, binary=False):
        '''
        file:
            type, string
            Path to file to iomanage
        start:
            -- OPTIONAL --
            type, boolean
            default, True
            Start operations thread on creation
        jtype:
            -- OPTIONAL --
            type, boolean
            default, True
            File is json database
        binary:
            -- OPTIONAL --
            type, boolean
            default, False
            Open file in binary read/write mode
        '''

        self.Ops = [] # Operations
        self.Out = {} # Outputs
        self.Reserved = [] # Reserved keys for operations

        self.stopthread = False # Should stop operations thread
        self.stopped = True # Is operations thread stopped
        self.thread = None # Operation thread object
        self.file = file # File to read/write

        ## Assigning open params to class

        if binary: # Can not be json type and binary read/write
            self.jtype = False
        else:
            self.jtype = jtype

        self.binary = binary

        # Create file if it doesn't already exist
        if not os.path.isfile(file):
            with open(file, "w+") as file:
                if jtype:
                    file.write("{}")

        if start: # start if kwarg start is True
            self.Start()

    def GetId(self): # Class to get a probably unique key
        return uuid.uuid4()

    def Read(self, waitforwrite=False, id=None): # Handles creating read operations
        '''
        waitforwrite:
            -- OPTIONAL --
            type, boolean
            default, False
            Operations thread should wait for write process same id kwarg
            Requires id kwarg to be set
        id:
            -- OPTIONAL --
            type, uuid4
            default, None
            ID to identify this operation
        '''

        if not waitforwrite:
            if id == None: id = uuid.uuid4() # get uuid if none passed
            self.Ops.append({"type": "r", "wfw": False, "id": id}) # Add operation to list
        else: # Wait for next write with same id
            if id == None: # waitforwrite requires id
                return None

            # Check for duplicate ids

            for x in self.Ops:
                if x["id"] == id:
                    return None

            if id in self.Reserved:
                return None

            # Reserve id
            # Add operation to list
            self.Reserved.append(id)
            self.Ops.append({"type": "r", "wfw": True, "id": id})

        while not id in self.Out: # Wait for read operation to complete
            time.sleep(.01)

        result = self.Out[id] # Get results
        del self.Out[id] # Delete results from output
        return result["data"] # return results

    def Write(self, nd, id=None):
        '''
        nd:
            type, string/bytes
            New data to write to file
        id:
            -- OPTIONAL --
            type, uuid
            default, None
            ID to identify this operation
        '''

        self.Ops.append({"type": "w", "d": nd, "id": id}) # Add write operation to list

    def Start(self): # Start operations thread
        if self.stopped: # Start only if thread not running
            self.stopthread = False # Reset stopthread to avoid immediate stoppage

            # Create thread and start
            self.thread = threading.Thread(target=self.ThreadFunc)
            self.thread.start()

    def Stop(self): # Stop operations thread
        if not self.stopthread: # Stop thread only if not already stopping
            if not self.stopped: # Stop thread only if thread running
                self.stopthread = True

    def isStopped(self): # Test if operations thread not running
        return self.stopped

    def ThreadFunc(self): # Operations function
        self.stopped = False # Reset stopped attr

        # Read/write type, binary or not
        t = None
        if self.binary:
            t = "b"
        else:
            t = ""

        # Main loop
        while not self.stopthread: # Test for stop attr
            if len(self.Ops) > 0: # Test for new operations

                # Get next operation
                Next = self.Ops[0]
                del self.Ops[0]

                # Open file as 'type' (read/write) + t (binary/text)
                with open(self.file, Next["type"]+t) as file:
                    id = Next["id"] # Operation ID

                    if Next["type"] == "r": # If is read operation

                        # Use json.load if in json mode
                        if self.jtype:
                            d = json.load(file)
                        else:
                            d = file.read()

                        # Put data in output
                        self.Out[id] = {"data": d, "id": id}

                        if Next["wfw"]: # Test if read operation is wait-for-write
                             # Wait for write loop
                            while not self.stopthread: # Test for stop attr

                                # Search for write operation with same id
                                op = None
                                for op in self.Ops:
                                    if op["id"] == id:
                                        break

                                # If no write operation, wait and restart loop
                                if op == None:
                                    time.sleep(.1)
                                    continue

                                self.Reserved.remove(id) # Remove reserved id
                                self.Ops.remove(op) # Remove write operation from list
                                self.Ops.insert(0, op) # Place write operation first
                                break # Break wfw loop
                            continue # Continue to main loop start

                    elif Next["type"] == "w": # If is write operation

                        # Use json.dump if in json mode
                        if self.jtype:
                            json.dump(Next["d"], file, indent=4)
                        else:
                            file.write(Next["d"])

            else: # If no operations, wait.
                time.sleep(.1)

        self.stopped = True # Set operation thread as stopped

class Manager:
    def __init__(self):
        self.gamePath = None
        self.R2API = None
        self.BIEP = None

        self.setupCache()

    def is_online(self):
        return requests.get("https://google.com").status_code == 200

    def is_64bit(self):
        return 'PROGRAMFILES(X86)' in os.environ

    def outdated(self, ov, nv):
        oParts = ov.split(".")
        nParts = nv.split(".")

        # Check if we have the version yet
        if ov.count("0") == len(oParts): return False

        if len(oParts) >= len(nParts):
            for i in range(0, len(oParts)):
                if int(nParts[i]) > int(oParts[i]):
                    return True
                elif int(nParts[i]) > int(oParts[i]):
                    return False
                if i+2 > len(nParts) and i+2 <= len(oParts):
                    return False
        else:
            for i in range(0, len(nParts)):
                if int(nParts[i]) > int(oParts[i]):
                    return True
                elif int(nParts[i]) > int(oParts[i]):
                    return False
                if i+2 > len(oParts) and i+2 <= len(nParts):
                    return True

    def install_biep(self):
        # Get Latest Version

        latest = requests.get("https://api.github.com/repos/BepInEx/BepInEx/releases").json()[0]["tag_name"][1:]

        print("Please wait, downloading BepInExPack v"+latest+"...")
        with open("BIEP.zip", "wb+") as f:
            if self.is_64bit:
                f.write(requests.get("https://github.com/BepInEx/BepInEx/releases/latest/download/BepInEx_x64_"+latest+".0.zip").content)
            else:
                f.write(requests.get("https://github.com/BepInEx/BepInEx/releases/latest/download/BepInEx_x86_"+latest+".0.zip").content)

        print("Extracting BIEP.zip...")
        with ZipFile("BIEP.zip", "r") as zO:
            zO.extractall("./BIEP")

        print("Merging BepInEx v"+latest+" with Risk of Rain 2...")
        for i in os.listdir("./BIEP"):
            try:
                shutil.move(os.path.join(os.getcwd()+"\\BIEP", i), self.gamePath)
            except Exception as e:
                print("Failed to move "+i+"\n"+str(e))

        print("Removing ./BIEP...")
        shutil.rmtree("./BIEP")

        print("Removing ./BIEP.zip...")
        os.remove("./BIEP.zip")

        self.BIEP = latest

    def update_biep(self):
        # Back everything up

        print("Backing up R2API...")
        os.mkdir("./R2API")

        for i in os.listdir(self.gamePath+"\\BepInEx"):
            if not i.endswith("core"):
                shutil.move(os.path.join(self.gamePath+"\\BepInEx", i), "./R2API")

        print("Removing old BIEP...")
        shutil.rmtree(self.gamePath+"\\BepInEx")
        os.remove(self.gamePath+"\\winhttp.dll")
        if os.path.isfile(self.gamePath+"\\changelog.txt"): os.remove(self.gamePath+"\\changelog.txt")

        self.install_biep()

        print("Restoring R2API...")
        for i in os.listdir("./R2API"):
            shutil.move(os.path.join(os.getcwd()+"\\R2API", i), self.gamePath+"\\BepInEx")
        shutil.rmtree("./R2API")

        print("BepInExPack has been updated!")

    def update_r2api(self):
        # Back up mods

        print("Backing up mods...")
        os.mkdir("./Mod-Backups")

        for i in os.listdir(self.gamePath+"\\BepInEx\\plugins"):
            if i != "R2API":
                shutil.move(os.path.join(self.gamePath+"\\BepInEx\\plugins", i), "./Mod-Backups")

        print("Removing old R2API...")
        shutil.rmtree(self.gamePath+"\\BepInEx\\plugins")
        shutil.rmtree(self.gamePath+"\\BepInEx\\monomod")
        os.remove(self.gamePath+"\\BepInEx\\icon.png")
        os.remove(self.gamePath+"\\BepInEx\\manifest.json")
        os.remove(self.gamePath+"\\BepInEx\\README.md")

        self.install_r2api()

        print("Restoring mods...")
        for i in os.listdir("./Mod-Backups"):
            shutil.move(os.path.join(os.getcwd()+"\\Mod-Backups", i), self.gamePath+"\\BepInEx\\plugins")

        shutil.rmtree("./Mod-Backups")

        print("R2API has been updated!")

    # Gets mods installed without this manager
    def get_current_mods(self):
        if os.path.isdir(self.gamePath+"\\BepInEx\\plugins"):
            id = self.configs.GetId()
            config = self.configs.Read(waitforwrite=True, id=id)
            for i in os.listdir(self.gamePath+"\\BepInEx\\plugins"):
                if not "R2API" in i:
                    with open(os.path.join(self.gamePath+"\\BepInEx\\plugins", i+"\\manifest.json"), "r") as file:
                        pc = json.load(file)

                    if not pc["name"] in config["cachedMods"]:
                        config[pc["name"]] = pc
            self.configs.Write(config, id)

    def install_mod(self, url):
        requirements = []

        details = str(requests.get(url).content)[2:]
        details = details[:len(details)-1]
        author = url.split("/package/")[1].split("/")[0]
        name = url.split("/package/")[1].split("/")[1]
        version = details.split("<td>Dependency string</td>\\n        <td>")[1].split("</td>")[0].split("-")[-1]

        print("\n\nGetting details for "+name+" v"+version+" install...")
        if os.path.isdir(self.gamePath+"\\BepInEx\\plugins\\"+name):
            print(name+" is already installed.")
            return

        downloadurl = url.split("/package/")[0]+"/package/"+"download/"+url.split("/package/")[1]+version

        for x in details.split("<div class=\"list-group-item flex-column align-items-start media\">"):
            if x == details.split("<div class=\"list-group-item flex-column align-items-start media\">")[0]: continue
            a, n = x.split("<a href=\"/package/")[1].split("</a>")[0].split("\">")[1].split("-")
            requirements.append({"author": a, "name": n})

        print("Downloading "+name+" v"+version+"...")
        with open(name+".zip", "wb+") as f:
            f.write(requests.get(downloadurl).content)

        print("Extracting "+name+".zip...")
        with ZipFile(name+".zip", "r") as zO:
            zO.extractall("./"+name)

        # Make sure dir structure is correct
        if not os.path.isdir("./"+name+"/plugins") and not os.path.isfile("./"+name+"/"+name+".dll"):
            if os.path.isfile("./"+name+"/"+name+"/"+name+".dll"):
                shutil.move("./"+name+"/"+name+"/"+name+".dll", "./"+name)
            elif os.path.isdir("./"+name+"/"+name+"/plugins"):
                for i in os.listdir("./"+name+"/"+name):
                    shutil.move("./"+name+"/"+name+"/"+i, "./"+name+"/")

        if os.path.isfile("./"+name+"/manifest.json"):
            with open("./"+name+"/manifest.json", "r+") as file:
                try:
                    file.seek(3)
                    config = json.load(file)
                    config["author"] = author
                    file.seek(0)
                    file.truncate(0)
                    json.dump(config, file)
                except:
                    file.seek(0)
                    config = json.load(file)
                    config["author"] = author
                    file.seek(0)
                    file.truncate(0)
                    json.dump(config, file)

        if os.path.isdir("./"+name+"/plugins"):
            print("Merging with "+self.gamePath+"/BepInEx...")

            for i in os.listdir("./"+name+"/"):
                if os.path.isfile("./"+name+"/"+i):
                    try:
                        shutil.move("./"+name+"/"+i, self.gamePath+"/BepInEx/")
                    except Exception as e:
                        print("Failed to move "+i+", "+str(e))
                elif os.path.isdir("./"+name+"/"+i):
                    if not os.path.isdir(self.gamePath+"/BepInEx/"+i): os.mkdir(self.gamePath+"/BepInEx/"+i)

                    for x in os.listdir("./"+name+"/"+i):
                        try:
                            shutil.move("./"+name+"/"+i+"/"+x, self.gamePath+"/BepInEx/"+i+"/"+x)
                        except Exception as e:
                            print("Failed to move "+i+"/"+x+", "+str(e))

            shutil.rmtree("./"+name)
        else:
            print("Merging with "+self.gamePath+"\\BepInEx\\plugins...")
            shutil.move(os.getcwd()+"\\"+name, self.gamePath+"\\BepInEx\\plugins")

        print("Clearing junk...")
        os.remove("./"+name+".zip")

        print("Installing requirements...")
        for req in requirements:
            if req["name"] == "BepInExPack": continue
            if os.path.isdir(self.gamePath+"\\BepInEx\\plugins\\"+req["name"]):
                print(req["name"]+" is already installed.")
            else:
                self.install_mod("https://thunderstore.io/package/"+req["author"]+"/"+req["name"]+"/")

        print(name+" v"+version+" has been successfully installed.")

    def launch_nw(self):

        #while True:

        online = self.is_online()

        if self.BIEP == None and online:
            if input("\n\nBepInEx is not installed! Would you like to install it (Required for mod use)? (y/n) ")[0].lower() == "y":
                self.install_biep()
        elif not self.BIEP == None:
            print("BIEP Install Version: v"+self.BIEP)

        if self.R2API == None and online:
            if input("\n\nR2API is not installed! Would you like to install it (Required for mod use)? (y/n) ")[0].lower() == "y":
                self.install_mod("https://thunderstore.io/package/tristanmcpherson/R2API/")
        elif not self.BIEP == None:
            print("R2API Install Version: v"+self.R2API)

        if self.outdated(self.BIEP, requests.get("https://api.github.com/repos/BepInEx/BepInEx/releases").json()[0]["tag_name"][1:]):
            if input("\n\nThere is a newer version of BepInExPack. Would you like to install it? (y/n) ")[0].lower() == "y":
                self.update_biep()

        if input("\n\nWould you like to install Kat's recommended mods?\nThese mods will have little affect on gameplay and are quality of life mods. (y/n) ")[0].lower() == "y":
            mods = ["https://thunderstore.io/package/Harb/DebugToolkit/", "https://thunderstore.io/package/JohnEdwa/RTAutoSprintEx/", "https://thunderstore.io/package/Lodington/Thiccify/", "https://thunderstore.io/package/DekuDesu/SkipWelcomeScreen/", "https://thunderstore.io/package/Kazzababe/SavedGames/", "https://thunderstore.io/package/xayfuu/EnemyHitLog/", "https://thunderstore.io/package/RyanPallesen/VanillaTweaks/", "https://thunderstore.io/package/Pickleses/TeleporterShow/", "https://thunderstore.io/package/mpawlowski/Compass/", "https://thunderstore.io/package/DekuDesu/MiniMapMod/", "https://thunderstore.io/package/SushiDev/DropinMultiplayer/", "https://thunderstore.io/package/pixeldesu/Pingprovements/", "https://thunderstore.io/package/IFixYourRoR2Mods/DiscordRichPresence/", "https://thunderstore.io/package/TheRealElysium/EmptyChestsBeGone/", "https://thunderstore.io/package/kookehs/StatsDisplay/"] # Broken "https://thunderstore.io/package/vis-eyth/UnmoddedClients/", "https://thunderstore.io/package/RyanPallesen/AssortedSkins/", https://thunderstore.io/package/felixire/BUT_IT_WAS_ME_DIO/
            for mod in mods:
                self.install_mod(mod)


    def getGamePath(self):
        if os.path.isfile("C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Steam/Steam.lnk"):
            steam = win32com.client.Dispatch("WScript.Shell").CreateShortCut("C:/ProgramData/Microsoft/Windows/Start Menu/Programs/Steam/Steam.lnk").Targetpath.split("\\")
            del steam[-1]
            steam = "/".join(steam)

            s_install_folders = [steam]

            with open(steam+"/config/config.vdf", "r") as file:
                try:
                    d = vdf.load(file)["InstallConfigStore"]["Software"]["valve"]["Steam"]
                except:
                    d = vdf.load(file)["InstallConfigStore"]["Software"]["Valve"]["Steam"]

                base_number = 1
                while True:
                    if "BaseInstallFolder_"+str(base_number) in d:
                        s_install_folders.append(d["BaseInstallFolder_"+str(base_number)])
                        base_number += 1
                    else:
                        break

                for folder in s_install_folders:
                    if os.path.isdir(folder+"/steamapps/common/Risk of Rain 2") and os.path.isfile(folder+"/steamapps/common/Risk of Rain 2/Risk of Rain 2.exe"):
                        print("Found RoR2 install automatically at "+folder+"/steamapps/common/Risk of Rain 2")
                        return folder+"/steamapps/common/Risk of Rain 2"

        w = QWidget()
        w.setWindowTitle("Select Risk Of Rain 2 Directory")
        w.show()

        return str(QFileDialog.getExistingDirectory(w, "Select Risk Of Rain 2 Directory"))

    def setupCache(self):
        # Setup cache files

        if not os.path.isdir("./Mods"):
            os.mkdir("./Mods")

        if not os.path.isfile("./configs.json"):
            with open("./configs.json", "w+"): pass

        with open("./configs.json", "r+") as file:
            if file.read() == "":
                file.seek(0)
                dc = {"gamePath": self.getGamePath(), "cachedMods": {}, "modProfiles": []}
                json.dump(dc, file, indent=4)
            else:
                file.seek(0)
                dc = json.load(file)
                if not os.path.isdir(dc["gamePath"]) or not\
                    os.path.isfile(dc["gamePath"]+"Risk Of Rain 2.exe" if\
                    dc["gamePath"].endswith("/") else dc["gamePath"]+"/Risk Of Rain 2.exe"):

                    dc["gamePath"] = self.getGamePath()
                    json.dump(dc, file, indent=4)

        self.gamePath = dc["gamePath"]
        self.configs = IOManager("./configs.json")

        # Are mod dependencies installed?

        if os.path.isdir(self.gamePath+"/BepInEx") and os.path.isfile(self.gamePath+"/winhttp.dll"):
            if os.path.isfile(self.gamePath+"/BepInEx/LogOutput.log"):
                with open(self.gamePath+"/BepInEx/LogOutput.log", "r") as f:
                    f = f.readline()
                    if "BepInEx" in f and "-" in f:
                        self.BIEP = f.split("BepInEx")[1].split("-")[0].strip()
                    else:
                        self.BIEP = "0.0.0.0"
            else:
                self.BIEP = "0.0.0.0"

            if os.path.isfile(self.gamePath+"/BepInEx/manifest.json") and os.path.isfile(self.gamePath+"/BepInEx/monomod/Assembly-CSharp.R2API.mm.dll"):
                with open(self.gamePath+"/BepInEx/manifest.json", "r+") as f:
                    try:
                        f.seek(3) # Avoid gunk at start of file
                        f = json.load(f)
                        self.R2API = f["version_number"]
                    except:
                        f.seek(0) # Avoid gunk at start of file
                        f = json.load(f)
                        self.R2API = f["version_number"]
            elif os.path.isfile(self.gamePath+"/BepInEx/monomod/Assembly-CSharp.R2API.mm.dll"):
                self.R2API = "0.0.0.0"

        # Get mods installed without this manager
        self.get_current_mods()

if __name__ == "__main__":
    qa = QApplication(sys.argv)

    m = Manager()
    #try:
    m.launch_nw()
    #except Exception as e:
    #    print("Error: "+str(e))

    m.configs.Stop()
