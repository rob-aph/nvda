import globalPluginHandler
from globalCommands import SCRCAT_SYSTEMCARET
from scriptHandler import script
import api
import speech
import controlTypes
import textInfos
import tones

class GlobalPlugin(globalPluginHandler.GlobalPlugin):
	def getTextInfoAtCaret(self):
		# returns None if not editable text
		ti = None
		focus = api.getFocusObject()
		if focus.role == controlTypes.Role.EDITABLETEXT:
			try:
				ti = focus.makeTextInfo(textInfos.POSITION_CARET)
			except:
				pass

		return ti

	def readCurrentLine(self, ti):
		# ti is TextInfo object
		tempTi = ti.copy()
		tempTi.expand(textInfos.UNIT_LINE)
		speech.speakMessage(tempTi.text)

	def moveToBlockParagraph(self, next):
		MAX_LINES = 250 # give up after searching this many lines
		ti = self.getTextInfoAtCaret()
		if ti is None: return False
		moved = False
		lookingForBlank = True
		moveOffset = 1 if next else -1
		lines = 0
		while lines < MAX_LINES:
			tempTi = ti.copy()
			tempTi.expand(textInfos.UNIT_LINE)
			isBlank = len(tempTi.text.strip()) == 0
			if lookingForBlank and isBlank:
				lookingForBlank = False
			if not lookingForBlank and not isBlank:
				moved = True
				break
			if not ti.move(textInfos.UNIT_LINE, moveOffset): break # next line
			lines += 1

		# exception: if moving backwards, need to move to top of now current paragraph
		if moved and not next:
			while True:
				if not ti.move(textInfos.UNIT_LINE, -1):
					break # leave at top
				tempTi = ti.copy()
				tempTi.expand(textInfos.UNIT_LINE)
				if not len(tempTi.text.strip()):
					# found blank line before desired paragraph
					ti.move(textInfos.UNIT_LINE, 1) # first line of paragraph
					break

		if moved:
			ti.updateCaret()
			self.readCurrentLine(ti)
		else:
			tones.beep(1000, 30)

		return True

	@script(
		# Translators: Input help mode message for next block paragraph command
		description=_("Moves to the next block paragraph."),
		category=SCRCAT_SYSTEMCARET,
		gesture="kb:control+downArrow"
	)
	def script_nextBlockParagraph(self, gesture):
		if not self.moveToBlockParagraph(True):
			gesture.send()

	@script(
		# Translators: Input help mode message for previous block paragraph command
		description=_("Moves to the previous block paragraph."),
		category=SCRCAT_SYSTEMCARET,
		gesture="kb:control+upArrow"
	)
	def script_previousBlockParagraph(self, gesture):
		if not self.moveToBlockParagraph(False):
			gesture.send()

