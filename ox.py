import xlwings as xw
import random
import re

PLAY_AREA = "A1:C3"

def msg(sht, msg):
  sht.range(5, 2).value = msg

@xw.sub
def my_macro():
  wb = xw.Book.caller()
  sht = wb.sheets[0]

  # ready 3x3 area
  area = sht.range(PLAY_AREA).value

  # check game over (1/2)
  winner = who_win(area)

  if winner is not None:
    msg(sht, "%s win!" % winner)
  else:
    # find blank slot
    slots = []
    for i in range(1, 4):
      for j in range(1, 4):
        v = sht.range(i, j).value
        if v is None:
          slots.append([i, j])

    if(len(slots) == 0):
    # check blank slot
      msg(sht, 'Slot is full, draw game!')
    else:
    # or random mark 'x'
      dest = random.choice(slots)
      msg(sht, "x marks %s" % dest)
      sht.range(dest[0], dest[1]).value = 'x'

      # check game over (2/2)
      area = sht.range(PLAY_AREA).value
      winner = who_win(area)
      if winner is not None:
        msg(sht, "%s win!" % winner)

def who_win(area):
  win_patt = [
    "xxx......",
    "...xxx...",
    "......xxx",
    "x..x..x..",
    ".x..x..x.",
    "..x..x..x",
    "x...x...x",
    "..x.x.x..",
  ]
  # check o win
  ooo = oneline('o', area)
  for idx, val in enumerate(win_patt):
    patt = re.compile(val)
    if patt.match(ooo):
      return 'o'
  # check x win
  xxx = oneline('x', area)
  for idx, val in enumerate(win_patt):
    patt = re.compile(val)
    if patt.match(xxx):
      return 'x'
  # game not end
  return None

def oneline(p, area):
  l = ''
  for i in range(0, 3):
    for j in range(0, 3):
      #v = sht.range(i, j).value
      v = area[i][j]
      x = 'x' if v == p else '-'
      l += x
  return l
