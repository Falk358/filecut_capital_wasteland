import re
import uno
import unohelper
import json
import operator
import math
import functools

from com.sun.star.datatransfer import XTransferable, DataFlavor
from com.sun.star.awt import MessageBoxButtons as MSG_BUTTONS
from com.sun.star.table import CellRangeAddress

def macro(func):
    """
    Decorator used for every macro.
    Ensure that the context is always updated and that the document is
    compatible before launching the macro. This way, macro bodies can focus
    on doing useful stuff instead of checks.
    """

    @functools.wraps(func)
    def wrapper_decorator(*args, **kwargs):

        self = args[0]

        self.update_context()

        if func.__name__ == "about":
            func(self)
        else:
            if self.document_compatible:
                func(self)
            else:
                self.msgbox("This document is not compatible with the Filecutter Toolkit", title="Compatibility problem")

    return wrapper_decorator

class FileCutterToolkit(object):

    def __init__(self):

        self.VERSION = "1.4"
        self.COMPATIBLE_SCRIPT_VERSIONS = [1]
        self.FORMAT = {
                    1: {"SPEAKER": 0, "NPCID": 1 , "RACE": 2, "VOICE TYPE": 3, "QUEST": 4, "CATEGORY": 5, "TOPIC": 6, "TOPICINFO": 7, "REPONSE INDEX": 8, "FILENAME": 9, 
                        "FULL PATH":10, "FILE FOUND": 11, "MODIFIED SINCE FILE CREATION": 12, "TOPIC TEXT": 13, "PROMPT": 14, "RESPONSE TEXT": 15, "EMOTION": 16, "SCRIPT NOTES": 17, "FILECUTTER NOTES": 18}
                  }

        # Colors definition
        self.PERFECT_COLOR =       0x00b0f0 # Cyan
        self.IB_COLOR =            0xffff00 # Yellow
        self.MISSING_COLOR =       0xff0000 # Red
        self.MISPELLED_COLOR =     0x7030a0 # Purple
        self.MISPRONUNCED_COLOR =  0x0070c0 # Blue
        self.BAD_ACTING_COLOR =    0x00b050 # Green
        self.SOUND_QUALITY_COLOR = 0xffc000 # Orange
        self.IGNORE_COLOR =        0x000000 # Black

        self.update_context()

    def update_context(self):
        """
        Update volatile data. Usefull to allow the use of the toolkit on
        multiple scripts at the same time
        Must be called first thing in any macro. (insured with macro decorator)
        """

        self.ctx = uno.getComponentContext()
        self.sm = self.ctx.getServiceManager()
        self.desktop = XSCRIPTCONTEXT.getDesktop()
        self.model = self.desktop.getCurrentComponent()
        self.sheet = self.model.CurrentController.ActiveSheet
        self.script_version = self.get_script_version()
        if self.script_version in self.COMPATIBLE_SCRIPT_VERSIONS:
            self.current_format = self.FORMAT[self.script_version]
            self.document_compatible = True
        else:
            self.document_compatible = False

    # -----------------------[ SCRIPT MANIPULATION ]----------------------------

    def get_script_version(self):
        """
        Returns the script version or -1 if this is not a VA script
        """
        cell_content = self.sheet.getCellByPosition(9, 0).String
        #self.msgbox(f"{cell_content}", "test")

        if cell_content == "FILENAME":
            return 1
        return -1

    def get_line_from_filename(self, filename):
        """
        Returns the number of the line that first match the filename, -1 if not found
        """

        start, end = self.get_script_limits()

        for i in range(start, end):
            if self.get_line_data(i)["filename"] == filename:
                return i

        return -1

    def insert_script_line(self, line_number, line_data):
        """
        Create a new line in the script and set data into it
        """

        # Creating the line
        cell_range = CellRangeAddress()
        cell_range.Sheet = 0
        cell_range.StartColumn = 0
        cell_range.EndColumn = len(self.current_format)
        cell_range.StartRow = line_number
        cell_range.EndRow = line_number
        self.sheet.insertCells(cell_range, 3)

        # Setting data
        self.set_line_data(line_number, line_data)

        # Setting color to default
        row_cells = []

        for i in range(len(self.current_format)):
            row_cells.append(self.sheet.getCellByPosition(i, line_number))

        for cell in row_cells:
            cell.IsCellBackgroundTransparent = True

    def set_line_data(self, line_number, line_data):
        """
        Set data into a line
        """

        row_cells = []

        for i in range(len(self.current_format)):
            row_cells.append(self.sheet.getCellByPosition(i, line_number))

        for key in self.current_format:
            row_cells[self.current_format[key]].String = line_data[key]

    def get_script_limits(self):
        """
        Returns line numbers corresponding to the first and last lines of the script
        """

        script_start = 1
        script_stop = script_start
        while(self.get_line_data(script_stop)["FILENAME"] != ""):
            script_stop += 1

        return script_start, script_stop

    def get_line_data(self, line_number):
        """
        Returns the content of a line in the form of a dict
        Keys depends on the script format version. See FORMAT for reference
        """

        line_data = dict()
        # Get the filename
        for key in self.current_format:
            line_data[key] = self.sheet.getCellByPosition(self.current_format[key], line_number).String

        return line_data

    def commit_line(self, row_number, color, comment, clipboard=False, bonus_IB=False):
        """
        Actually takes action on a line using the provided informations.
        """

        total_cell = len(self.current_format)
        filename_cell = self.current_format["FILENAME"]
        comment_cell = self.current_format["FILECUTTER NOTES"]

        row_cells = []

        for i in range(total_cell):
            row_cells.append(self.sheet.getCellByPosition(i, row_number))

        for cell in row_cells:
            cell.CellBackColor = color

        row_cells[comment_cell].String = comment

        if clipboard:
            self.copy_to_clipboard(row_cells[filename_cell].String)
        else:
            self.copy_to_clipboard("")

    def get_line_color(self, line_number):
        """
        Return the line color in hex form
        """

        filename_cell = self.current_format["FILENAME"]
        cell = self.sheet.getCellByPosition(filename_cell, line_number)

        if cell.IsCellBackgroundTransparent:
            return 0xffffff
        else:
            return cell.CellBackColor

    # ----------------------[ LIBREOFFICE UTILITIES ]---------------------------

    def create_instance(self, name, with_context=False):

        if with_context:
            instance = self.sm.createInstanceWithContext(name, self.ctx)
        else:
            instance = self.sm.createInstance(name)
        return instance

    def msgbox(self, message, title, buttons=MSG_BUTTONS.BUTTONS_OK, type_msg='infobox'):
        """
        Create a messagebox in LibreOffice
        """

        toolkit = self.create_instance('com.sun.star.awt.Toolkit')
        parent = toolkit.getDesktopWindow()
        mb = toolkit.createMessageBox(parent, type_msg, buttons, "Filecutter Toolkit - {}".format(title), str(message))
        return mb.execute()

    def copy_to_clipboard(self, data):
        """
        Get that data into the clipboard for labeling.
        """

        transferable = Transferable(data)
        oClip = self.ctx.getServiceManager().createInstanceWithContext(
            "com.sun.star.datatransfer.clipboard.SystemClipboard", self.ctx)
        oClip.setContents(transferable, None)

    # -------------------------[ STATS UTILITIES ]------------------------------

    def quality_report_percent(self, part, whole):
        if whole == 0:
            return 0
        else:
            return math.floor(float(part) / whole * 100)

    # -----------------------------[ MACROS ]-----------------------------------

    @macro
    def about(self):
        """
        Print a messagebox with various informations about the Filecutter Toolkit
        """

        message = "Filecutter Toolkit v{}\n".format(self.VERSION)
        if self.script_version != -1:
            message += "Detected script format: v{}\n".format(self.script_version)
            if self.script_version in self.COMPATIBLE_SCRIPT_VERSIONS:
                message += "Compatible with script format: YES"
            else:
                message += "Compatible with script format: NO"
        else:
            message += "This document does not appear to be a valid voice acting script"
        self.msgbox(message, title="About")

    @macro
    def perfect(self):
        """
        Apply the right color and comment for a good line based on the filename.
        The filename is put into the clipboard for labeling purpose.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        self.commit_line(row_number, self.PERFECT_COLOR, "Perfect", clipboard=True)

    @macro
    def mispelled(self):
        """
        Apply the right color and template comment for a mispelled line.
        The filename is put into the clipboard for labeling purpose as this is not
        a VA defect.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        comment = "Script error: TODO red highlight of mispell and comment"
        self.commit_line(row_number, self.MISPELLED_COLOR, comment, clipboard=True)

    @macro
    def sound_quality(self):
        """
        Apply the right color and template comment for a line that is not up to
        standard quality wise.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        comment = "Sound quality: TODO describe the problem"
        self.commit_line(row_number, self.SOUND_QUALITY_COLOR, comment)

    @macro
    def bad_acting(self):
        """
        Apply the right color and template comment for a line that is not up to
        standard acting wise.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        comment = "Acting: TODO helpful comment for the voice actor"
        self.commit_line(row_number, self.BAD_ACTING_COLOR, comment)

    @macro
    def mispronunced(self):
        """
        Apply the right color and template comment for a line that is mispronunced.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        comment = "Mispronunciation: TODO red highlight of mispronunced word and comment"
        self.commit_line(row_number, self.MISPRONUNCED_COLOR, comment)

    @macro
    def missing(self):
        """
        Apply the right color and comment for a missing line.
        """

        row_number = self.model.CurrentSelection.RangeAddress.StartRow
        line_data = self.get_line_data(row_number)

        comment = "Missing"
        self.commit_line(row_number, self.MISSING_COLOR, comment)

    @macro
    def statistics(self):
        """
        Prints a messagebox with stats about the filecutting process.
        """

        start, end = self.get_script_limits()

        total = end - start
        perfect = 0
        mispelled = 0
        sound_quality = 0
        bad_acting = 0
        mispronunced = 0
        missing = 0
        ignored = 0
        todo = 0
        template_left = 0

        for i in range(start, end):
            color = self.get_line_color(i)
            filecutter_notes = self.get_line_data(i)["FILECUTTER NOTES"]

            if "TODO " in filecutter_notes:
                template_left += 1

            if color == self.PERFECT_COLOR:
                perfect += 1
            elif color == self.MISPELLED_COLOR:
                mispelled += 1
            elif color == self.SOUND_QUALITY_COLOR:
                sound_quality += 1
            elif color == self.BAD_ACTING_COLOR:
                bad_acting += 1
            elif color == self.MISPRONUNCED_COLOR:
                mispronunced += 1
            elif color == self.MISSING_COLOR:
                missing += 1
            elif color == self.IGNORE_COLOR:
                ignored += 1
            else:
                todo += 1

        already_cut =  perfect + mispelled + sound_quality + bad_acting + mispronunced + missing
        advancement = math.floor((1 - (float(todo) / total)) * 100)

        message = """
        Progress report ---------
        Progress:         {}%
        Total lines:       {}
        Ignored lines:  {}
        Already cut :    {}
        Yet to cut:        {}

        Cutting quality report --
        Forgotten TODO:  {}

        VA quality report -------
        Perfect:               {} ({}%)
        Mispelled:          {} ({}%)
        Sound Quality:  {} ({}%)
        Bad Acting:        {} ({}%)
        Mispronunced:  {} ({}%)
        Missing:              {} ({}%)
        """.format(advancement, total, ignored, already_cut, todo, template_left, perfect, self.quality_report_percent(perfect, already_cut), mispelled, self.quality_report_percent(mispelled, already_cut), sound_quality, self.quality_report_percent(sound_quality, already_cut), bad_acting, self.quality_report_percent(bad_acting, already_cut), mispronunced, self.quality_report_percent(mispronunced, already_cut), missing, self.quality_report_percent(missing, already_cut))

        self.msgbox(message, title="Statistics")

class Transferable(unohelper.Base, XTransferable):
    """Keep clipboard data and provide them."""

    def __init__(self, text):
        df = DataFlavor()
        df.MimeType = "text/plain;charset=utf-16"
        df.HumanPresentableName = "encoded text utf-16"
        self.flavors = [df]
        self.data = [text] #[text.encode('ascii')]

    def getTransferData(self, flavor):
        if not flavor: return
        mtype = flavor.MimeType
        for i,f in enumerate(self.flavors):
            if mtype == f.MimeType:
                return self.data[i]

    def getTransferDataFlavors(self):
        return tuple(self.flavors)

    def isDataFlavorSupported(self, flavor):
        if not flavor: return False
        mtype = flavor.MimeType
        for f in self.flavors:
            if mtype == f.MimeType:
                return True
        return False

# ------------------------[ PUBLIC MACRO EXPORT ]-------------------------------

TOOLKIT = FileCutterToolkit()

About = TOOLKIT.about
Perfect = TOOLKIT.perfect
Mispelled = TOOLKIT.mispelled
SoundQuality = TOOLKIT.sound_quality
BadActing = TOOLKIT.bad_acting
Mispronunced = TOOLKIT.mispronunced
Missing = TOOLKIT.missing
Statistics = TOOLKIT.statistics

g_exportedScripts = (About, Perfect, Mispelled, SoundQuality, BadActing, Mispronunced, Missing, Statistics)
