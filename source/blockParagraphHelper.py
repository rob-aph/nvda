import api
import speech
import controlTypes
import textInfos
import tones
from NVDAObjects.IAccessible.winword import WordDocument as IAccessibleWordDocument
from NVDAObjects.UIA.wordDocument import WordDocument as UIAWordDocument

def getTextInfoAtCaret():
	# returns None if not editable text or if in Microsoft Word document
	ti = None
	focus = api.getFocusObject()
	if isinstance(focus, IAccessibleWordDocument) or isinstance(focus, UIAWordDocument):
		return ti
	if not controlTypes.State.MULTILINE in focus.states:
		return ti
	if focus.role == controlTypes.Role.EDITABLETEXT:
		try:
			ti = focus.makeTextInfo(textInfos.POSITION_CARET)
		except (NotImplementedError, RuntimeError):
			pass

	return ti

def readCurrentParagraph(ti):
	# ti is TextInfo object
	paragraph = ""
	tempTi = ti.copy()
	while True:
		tempTi.expand(textInfos.UNIT_LINE)
		line = tempTi.text.strip()
		tempTi.collapse()
		if not len(line):
			break
		paragraph += line + "\r\n"
		if not tempTi.move(textInfos.UNIT_LINE, 1):
			break

	speech.speakMessage(paragraph)

def moveToBlockParagraph(next):
	MAX_LINES = 250  # give up after searching this many lines
	ti = getTextInfoAtCaret()
	if ti is None:
		return False
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
		if not ti.move(textInfos.UNIT_LINE, moveOffset):
			break  # next line
		lines += 1

	# exception: if moving backwards, need to move to top of now current paragraph
	if moved and not next:
		while True:
			if not ti.move(textInfos.UNIT_LINE, -1):
				break  # leave at top
			tempTi = ti.copy()
			tempTi.expand(textInfos.UNIT_LINE)
			if not len(tempTi.text.strip()):
				# found blank line before desired paragraph
				ti.move(textInfos.UNIT_LINE, 1)  # first line of paragraph
				break

	if moved:
		ti.updateCaret()
		readCurrentParagraph(ti)
	else:
		tones.beep(1000, 30)

	return True
