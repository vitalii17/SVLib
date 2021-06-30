import os
import pathlib
import sys
import configparser as config

def main(args):
    path             = ""
    script_directory = os.path.dirname(os.path.realpath(__file__))
    dir_files        = os.listdir(script_directory)
    if len(args) > 1:
        path = args[1]
    else:
        for item in dir_files:
            file_extension = pathlib.Path(item).suffix
            if file_extension == ".PrjPcb":
                path = item

    if path == "":
        sys.exit(2)
    
    prj_cfg = config.ConfigParser()
    prj_cfg.read(path)
    
    prj_title   = ""
    prj_version = ""
    
    for section in prj_cfg.sections():
        if prj_cfg.items(section)[0][1] == "ProjectTitle":
            prj_title = prj_cfg.items(section)[1][1]
        if prj_cfg.items(section)[0][1] == "Version":
            prj_version = prj_cfg.items(section)[1][1]
    
    # os.rename("a.txt", "b.txt")
    os.system("ren \"*.gtl\" \"%s-%s-Top.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gbl\" \"%s-%s-Bot.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gbs\" \"%s-%s-TopMask.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gbs\" \"%s-%s-BotMask.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gto\" \"%s-%s-TopSilk.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gbo\" \"%s-%s-BotSilk.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gtp\" \"%s-%s-TopPast.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*.gbp\" \"%s-%s-BotPast.gbr\"" % (prj_title, prj_version))
    os.system("ren \"*-Plated.txt\" \"%s-%s-Plated.drl\"" % (prj_title, prj_version))
    os.system("ren \"*-NonPlated.txt\" \"%s-%s-NonPlated.drl\"" % (prj_title, prj_version))
    
    print(prj_title)
    print(prj_version)
    
    # From old batch file:
    # set prj_name=Hive
    # set ver=1.0
    
    # ren "*.gbl" "%prj_name%-%ver%-Bot.gbr"
    # ren "*.gtl" "%prj_name%-%ver%-Top.gbr"
    # ren "*.gts" "%prj_name%-%ver%-TopMask.gbr"
    # ren "*.gbs" "%prj_name%-%ver%-BotMask.gbr"
    # ren "*.gm7" "%prj_name%-%ver%-Board.gbr"
    # ren "*.gto" "%prj_name%-%ver%-TopSilk.gbr"
    # ren "*.gbo" "%prj_name%-%ver%-BotSilk.gbr"
    # ren "*.gtp" "%prj_name%-%ver%-TopPast.gbr"
    # ren "*.gbp" "%prj_name%-%ver%-BotPast.gbr"
    # ren "*-Plated.txt" "%prj_name%-%ver%-Plated.drl"
    # ren "*NonPlated.txt" "%prj_name%-%ver%-NonPlated.drl"

if __name__ == '__main__':
    main(sys.argv)
