#! /usr/bin/env /usr/bin/python3
# for Windows change above to: hashbang python3

# Copyright (C) 2008-2010 by James C. Ahlstrom, N2ADR.
# Copyright (C) 2019 by Christopher Sylvain, KB3CS.
# This free software is licensed for use under the GNU General Public
# License (GPL), see http://www.opensource.org.
# Note that there is NO WARRANTY AT ALL.  USE AT YOUR OWN RISK!!

# Thanks to Chris, KB3CS, for additional code and features.
#
# adjusted SWR Threshold radiobuttons implementation            KB3CS 02/20/2010
# refactored for Python3.7  	                                KB3CS 09/04/2019
#   dependencies: pyserial
#   Win10 dependencies: pywin32 (formerly win32all)

# This controls the AT-200PC from LDG Electronics.  The buttons are:
#  Ant 1/2				Change antenna
#  Active/Passive		Passive: Zero added L and C, turn off AutoTune (pass thru)
#  Auto On/Off			Turn AutoTune on and off (auto start tune for high SWR)
#  Mem Tune				Tune from memory
#  L, C +/-				Add / Subtract one from inductance L or capacitance C
#  Z					Change direction of L-network
#  Store				Manually store L/C/Z for last frequency
#  Full Tune			Start a full tuning search
#  1.1, 1.3, 1.5, 1.7, 2.0, 2.5, 3.0
#                       Set / Show SWR autotuning threshold
# Please see the documentation for the AT-200PC or AT-200 Pro to understand this tuner.

import sys, time, math, traceback

import tkinter		# This is the standard tkinter GUI module.
from tkinter import messagebox

from types import *

DEBUG = 0

# This is the serial port name.  You probably need to change it:
if sys.platform[0:3] == "win":
  TTY_NAME = "COM4"		# Windows name of serial port for the AT-200PC
else:
  TTY_NAME = "/dev/ttyUSB0"	# Linux name of serial port for the AT-200PC

# These are other AT-200PC parameters you can set:
REQ_LIVEUPDATE = 63			# Send power and swr when RF is present: ON=63, OFF=64

# These are internal program parameters:
#FONT = 'helvetica -14 bold'	# Font for the screen text (not for buttons)
FONT = 'helvetica 13'
POLL_SEC = 0.05				# Time in seconds to poll the serial port
MAX_POWER = 200.0			# Power in watts for 100% scale
MAX_SWR = 5.0				# SWR for 100% scale (at least 1.1)

# text format prototypes using "fattest" digits
POWER_FMT_PROTO = "Forw  200W"
STATUS_FMT_PROTO = "Freq 20000 kHz,   Induc 100,   Capac 100,   High Z,   Tune Lost RF"

if sys.platform[0:3] == "win":
  try:
    import win32file		# This is part of the win32all module by Mark Hammond
  except:
    win32file = None
else:
  win32file = 1				# win32file not needed on Linux

try:
  import serial			# This is pySerial; it provides serial port support.
except:
  serial = None


def GetTextExtent(window, font, text):
  id = window.create_text(0, 0, text=text, font=font, anchor='nw')
  x1, y1, x2, y2 = window.bbox(id)
  w = x2 - x1
  h = y2 - y1
  window.delete(id)
  return w, h


class BaseButton:
  def __call__(self):
    if self.command:
      self.command(self)
  def GetValue(self):
    return self.var.get()
  def Display(self, value):
    self.var.set(value)
  def Nothing(self, event):	# Defeat change in color when mouse passes over
    return "break"


class BasePushbutton(tkinter.Button, BaseButton):
  def __init__(self, master, command, **kwd):
    self.command = command
    conf = {'master': master, 'command': self, 'text': 'N/A',
            'bd': 4, 'padx': 0, 'pady': 1, 'highlightthickness': 0,
            'width': 10,
        #'bg':'#A1A8DA',
        #    'disabledforeground':'#444',
        }
    conf.update(kwd)
    tkinter.Button.__init__(self, **conf)


class BaseCheckbutton(tkinter.Checkbutton, BaseButton):
  c_pushed = '#6F6'		# Color when pushed in
  def __init__(self, master, command, **kwd):
    self.command = command
    self.var = tkinter.IntVar()
    self.var.set(0)
    conf = {'master': master, 'command': self, 'text': 'N/A', 'indicatoron': 0,
            'bd': 4, 'padx': 0, 'pady': 2, 'variable': self.var,
            'highlightthickness': 0,
            'width': 10,
        #'bg':'#A1A8DA',				# normal background
        #'activebackground':'#0000FF',
        'selectcolor': self.c_pushed,		# Color when pushed in
        #'disabledforeground':'#444',
        }
    conf.update(kwd)
    tkinter.Checkbutton.__init__(self, **conf)
    self.c_gray = self.cget('bg')
    self.config(activebackground=self.c_gray)
  def __call__(self):
    if self.var.get():
      self.config(activebackground=self.c_pushed)
    else:
      self.config(activebackground=self.c_gray)
    BaseButton.__call__(self)


class BaseRadioButtons(BaseButton):		# A row of radio buttons
  c_pushed = '#6F6'		# Color when pushed in
  def __init__(self, master, command, labels, default, expand, **kwd):
    self.command = command
    self.expand = expand
    self.buttons = {}
    self.button_list = []
    # Determine the type of the data from the type of the first label
    if isinstance(labels[0], int):
      self.var = tkinter.IntVar()
    elif isinstance(labels[0], float):
      self.var = tkinter.DoubleVar();
    else:
      self.var = tkinter.StringVar()
    self.var.set(default)
    conf = {'master': master, 'command': self, 'indicatoron': 0,
            'bd': 3, 'padx': 0, 'pady': 2, 'variable': self.var, # 'bg':c_btn,
            'selectcolor': self.c_pushed,		# Color when pushed in
            'highlightthickness': 0,
            'disabledforeground': '#444',
           }
    if expand:		# Use expand=1, fill='x'
      conf['width'] = 1
    conf.update(kwd)
    if type(labels[0]) in (tuple, list):	# multiple rows
      for row in labels:
        frm = tkinter.Frame(master=master, bd=0, bg=c_bg)
        frm.pack(side='top', expand=1, fill='x')
        conf['master'] = frm
        self._AddRow(row, conf)
    else:
      self._AddRow(labels, conf)
    if default is not None:
      b = self.buttons[default]	# Currently selected
      b.config(activebackground=b.cget('selectcolor'))
  def _AddRow(self, row, conf):
      for itm in row:
        conf['value'] = v = itm
        conf['text'] = t = str(itm)
        b = tkinter.Radiobutton(**conf)
        b.the_value = v
        self.buttons[v] = b
        self.button_list.append(b)
        b.config(activebackground=b.cget('bg'))  # Turn off active background
        # This binding can defeat the button press!
        #b.bind(sequence='<Enter>', func=self.Nothing, add=0)
        #b.bind(sequence='<Leave>', func=self.Nothing, add=0)
        if not t:
          b.config(state='disabled')
        if self.expand:
          cf = {'expand':1, 'fill':'x'}
        else:
          cf = {}
        mx = 2	# was 6
        if itm == row[-1]:
          b.pack(side='left', anchor='w', padx=mx, pady=mx, **cf)
        else:
          b.pack(side='left', anchor='w', padx=(mx, 0), pady=mx, **cf)
  def __call__(self):
    self.command(self)
  def DisplayIndex(self, index):
    btn = self.button_list[index]
    var = btn.the_value
    self.var.set(var)
  def GetIndex(self):
    var = self.var.get()
    btn = self.buttons[var]
    return self.button_list.index(btn)


class Application(tkinter.Tk):
  def __init__(self):		# Draw all widgets
    tkinter.Tk.__init__(self)
    self.win_title = "AT200PC v1.4A on %s" % TTY_NAME
    self.wm_title(self.win_title)
    self.wm_resizable(0, 0)
    self.wm_protocol("WM_DELETE_WINDOW", self.WmDeleteWindow)
    self.wm_protocol("WM_SAVE_YOURSELF", self.WmDeleteWindow)
    fill = '#000'
    self.rx_state = 0
    self.serial = None
    self.tune_status = 'None'
    self.param1 = [None] * 20	# Parameters returned by the AT-200PC
    self.param2 = [None] * 20

    # create a top-level menu
    id = self.winfo_toplevel()
    m = tkinter.Menu(id, font=FONT, tearoff=0)
    m.add_command(label='Exit', underline=1, command=self.WmDeleteWindow) # Alt+X
    m.add_command(label='About..', underline=0, command=self.About) # Alt+A
    id.configure( menu=m )
  
    # create row of SWR Threshold buttons
    frm = tkinter.Frame(master=self, bd=2, relief='groove')
    frm.pack(side='bottom', anchor='s', expand=1, fill='both')
    id = tkinter.Label(frm, font=FONT, anchor='w', text='SWR Threshold')
    id.pack(side='left', pady=6, padx=8, fill='y')
    labels = (1.1, 1.3, 1.5, 1.7, 2.0, 2.5, 3.0)	# req 50 thru 56
    self.swrButns = BaseRadioButtons(frm, self.OnButtonSwr, labels, None, 1)
    
    # Create a row of buttons
    frm = tkinter.Frame(master=self, bd=2, relief='groove')
    frm.pack(side='bottom', anchor='s', expand=1, fill='both')
    id = tkinter.Label(frm, font=FONT, anchor='w', text=' ')
    id.pack(side='left', fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='L+', width=3)
    b.req = 1
    b.pack(side='left', anchor='w', padx=4, fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='L-', width=3)
    b.req = 2
    b.pack(side='left', anchor='w', padx=4, fill='y')
    id = tkinter.Label(frm, font=FONT, anchor='w', text=' ')
    id.pack(side='left', fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='C+', width=3)
    b.req = 3
    b.pack(side='left', anchor='w', padx=4, fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='C-', width=3)
    b.req = 4
    b.pack(side='left', anchor='w', padx=4, fill='y')
    id = tkinter.Label(frm, font=FONT, anchor='w', text=' ')
    id.pack(side='left', fill='y')
    b = BasePushbutton(frm, self.OnButtonHiLoZ, text='Z', width=3)
    b.pack(side='left', anchor='w', padx=4, fill='y')
    id = tkinter.Label(frm, font=FONT, anchor='w', text=' ')
    id.pack(side='left', fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='Store')
    b.req = 46
    b.pack(side='left', anchor='w', padx=4, fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='Full Tune')
    b.req = 6
    b.pack(side='right', anchor='e', padx=4, fill='y')

    scalex = b.winfo_reqwidth() * 5

    # Create a canvas for the meters and status
    Canvas = self.Canvas = tkinter.Canvas(self, bg='#FFF', bd=2, relief='groove')
    Canvas.pack(side='top')
    charx, chary = GetTextExtent(Canvas, FONT, '0')
    margin = max(2, chary / 4)
    h = max(5, chary / 10 * 7)		# height of swr/power bars
    dy = max(0, (chary - h) / 2)
    tab2 = GetTextExtent(Canvas, FONT, POWER_FMT_PROTO)[0] + margin * 2
    w, h = GetTextExtent(Canvas, FONT, STATUS_FMT_PROTO)

    scalex = max(scalex, w + margin * 2)

    # Create the SWR display
    id = Canvas.create_text(margin, margin, font=FONT, fill=fill, anchor='nw',
             text='SWR')
    self.swr = id
    x1, y1, x2, y2 = Canvas.bbox(id)
    y = y2
    id = Canvas.create_rectangle(tab2, y1 + dy, scalex, y2 - dy,
             width=1, outline = '#000')
    x1, y1, x2, y2 = Canvas.bbox(id)
    x1 = x1 + 1
    y1 = y1 + 1
    x2 = x2 - 1
    y2 = y2 - 1
    id = Canvas.create_rectangle(x1, y1, x2, y2, fill = '#0000FF')
    self.swr_meter = [id, x1, y1, x2, y2]
    self.swr_meter_size = x2 - x1

    # Create the Forward power display
    id = Canvas.create_text(margin, y, font=FONT, fill=fill, anchor='nw',
             text='Forw')
    self.power = id
    x1, y1, x2, y2 = Canvas.bbox(id)
    y = y2
    id = Canvas.create_rectangle(tab2, y1 + dy, scalex, y2 - dy,
             width=1, outline = '#000')
    x1, y1, x2, y2 = Canvas.bbox(id)
    x1 = x1 + 1
    y1 = y1 + 1
    x2 = x2 - 1
    y2 = y2 - 1
    id = Canvas.create_rectangle(x1, y1, x2, y2, fill = '#00FF00')
    self.power_meter = [id, x1, y1, x2, y2]
    self.power_meter_size = x2 - x1

    # Create the Reflected power display
    id = Canvas.create_text(margin, y, font=FONT, fill=fill, anchor='nw',
                            text='Refl')
    self.refl = id
    x1, y1, x2, y2 = Canvas.bbox(id)
    y = y2
    id = Canvas.create_rectangle(tab2, y1 + dy, scalex, y2 - dy,
                                 width=1, outline = '#000')
    x1, y1, x2, y2 = Canvas.bbox(id)
    x1 = x1 + 1
    y1 = y1 + 1
    x2 = x2 - 1
    y2 = y2 - 1
    id = Canvas.create_rectangle(x1, y1, x2, y2, fill = '#FFD700')
    self.refl_meter = [id, x1, y1, x2, y2]
    self.refl_meter_size = x2 - x1

    # Create the status text
    id = Canvas.create_text(margin, y, font=FONT, fill=fill, anchor='nw')
    if not win32file:
      Canvas.itemconfig(id, text="Missing Python module win32all")
    elif not serial:
      Canvas.itemconfig(id, text="Missing Python module pySerial")
    else:
      Canvas.itemconfig(id, text="Can not open serial port %s" % TTY_NAME)
    self.status1 = id
    x1, y1, x2, y2 = Canvas.bbox(id)
    Canvas.config(height=y2 + margin)
    Canvas.config(width=scalex + margin)

    # Create a row of buttons
    frm = tkinter.Frame(master=self, bd=2, relief='groove')
    frm.pack(side='bottom', anchor='s', expand=1, fill='both')
    b = self.antenna = BaseCheckbutton(frm, self.OnButtonAnt, text='Ant ?')
    b.pack(side='left', anchor='w', padx=4, fill='y')
    id = tkinter.Label(frm, font=FONT, anchor='w', text=' ')
    id.pack(side='left', fill='y')
    b = self.standby = BaseCheckbutton(frm, self.OnButtonStandby, text='Active')
    b.pack(side='left', anchor='w', padx=4, fill='y')
    b = self.btntune = BaseCheckbutton(frm, self.OnButtonAuto, text='Auto ON')
    b.pack(side='left', anchor='w', padx=4, fill='y')
    b = BasePushbutton(frm, self.OnButtonReq, text='Mem Tune')
    b.req = 5
    b.pack(side='right', anchor='e', padx=4, fill='y')

    self.running = 1

  def main(self):
    # Open the serial port, waiting if necessary.
    while not self.serial and self.running:
      if serial and win32file:
        try:
          self.serial = serial.Serial(port=TTY_NAME, timeout=0.05)
          self.serial.setRTS(0)			# turn off the RTS pin on the serial interface
        except serial.SerialException:
          pass
      time.sleep(0.1)
      self.update()
    if not self.serial:
      return
    # Send our requested initial state, and receive the AT-200PC state.
    # Wait for the AT-200PC to reply.
    self.Canvas.itemconfig(self.status1, text="Waiting for AT-200PC on %s" % TTY_NAME)
    time0 = 0.0
    while self.running:		# Send requested state plus REQ_ALLUPDATE (40)
      if 64 - REQ_LIVEUPDATE != self.param1[19]:
        self.Write(chr(REQ_LIVEUPDATE))
      elif self.param2[6] is None:		# We are assuming that param 6 is sent last
        if time.time() - time0 >= 1.0:	# Wait one second between requests
          time0 = time.time()
          self.Write(chr(40))		# Request an update
      elif self.param1[11] != 1:
        self.Write(chr(41))		# Request version
      else:
        break
      self.Read()			# Receive the current state of the AT-200PC
      self.update()
    if not self.running:
      return
    # Correct state has been received
    self.autotune = self.param1[17]
    self.NewData()		# Correct our controls for current state
    self.update()
    time0 = 0.0
    while self.running:
      if self.param2[6] is None:	# Send REQ_ALLUPDATE
        if time.time() - time0 >= 1.0:	# Wait one second between requests
          time0 = time.time()
          self.Write(chr(40))
      elif self.is_standby and self.param1[17] != 0:	# Standby implies no AutoTune
        self.Write(chr(59))				# Set autotune OFF
      elif not self.is_standby and self.autotune != self.param1[17]:
        self.Write(chr(59 - self.autotune))	# Set user's desired AutoTune state
      if self.Read():
        self.NewData()
      self.update()

  def WmDeleteWindow(self):
    if self.serial:
      self.serial.close()
      self.serial = None
    self.destroy()
    self.running = 0

  def OnButtonReq(self, btn):
    self.Write(chr(btn.req))

  def OnButtonHiLoZ(self, btn):
    if self.param1[3]:		# Currently Low impedance
      self.Write(chr(8))
    else:
      self.Write(chr(9))

  def OnButtonAnt(self, btn):
    if btn.GetValue():
      self.Write(chr(11))
    else:
      self.Write(chr(10))

  def OnButtonStandby(self, btn):
    if btn.GetValue():
      self.Write(chr(44))
    else:
      self.Write(chr(45))
    self.param2[6] = None	# Request relay settings

  def OnButtonAuto(self, btn):
    if btn.GetValue():
      self.autotune = 0
    else:
      self.autotune = 1

  def OnButtonSwr(self, btn):
    self.Write(chr(btn.GetIndex() + 50))

  def Write(self, s):		# Write a command string to the AT-200PC
    if DEBUG:
      print('Send', ord(s[0]))
    if self.serial:
      try:
        self.serial.setRTS(1)	# Wake up the AT-200PC
        time.sleep(0.003)		# Wait 3 milliseconds
        self.serial.write(s.encode())
        self.serial.setRTS(0)
        time.sleep(0.010)		# Wait
      except:
        traceback.print_exc()

  def Read(self):	# Receive characters from the AT-200PC
    change = 0	# Have any complete data blocks been received?
    if self.serial:
      try:
        chars = self.serial.read(1024)	# This will always time out
      except:
        chars = ''
        traceback.print_exc()
    else:
      chars = ''
    for ch in chars:
      if self.rx_state == 0:	# Read first of 4 characters; must be decimal 165
        if ch == 165:
          self.rx_state = 1
      elif self.rx_state == 1:	# Read second byte
        self.rx_state = 2
        self.rx_byte1 = ch
      elif self.rx_state == 2:	# Read third byte
        self.rx_state = 3
        self.rx_byte2 = ch
      elif self.rx_state == 3:	# Read fourth byte
        self.rx_state = 0
        byte3 = ch
        byte1 = self.rx_byte1
        byte2 = self.rx_byte2
        if DEBUG:
          print('Received', byte1, byte2, byte3)
        if byte1 > 19:	# Impossible command value
          continue
        if byte1 == 9:				# Tune pass
          self.tune_status = "OK"
          self.param2[6] = None		# Request relay settings
        elif byte1 == 10:			# Tune fail
          if byte2 == 0:
            self.tune_status = "No RF"
          elif byte2 == 1:
            self.tune_status = "Lost RF"
          elif byte2 == 2:
            self.tune_status = "High SWR"
          else:
            self.tune_status = "Error"
          self.param2[6] = None		# Request relay settings
        elif byte1 == 13:			# Start standby
          self.is_standby = 1
        elif byte1 == 14:			# Start active
          self.is_standby = 0
        self.param1[byte1] = byte2
        self.param2[byte1] = byte3
        change = 1
    return change

  def NewData(self):	# Change screen to show new data
    # Set Forward power display
    power = (self.param1[5] * 256 + self.param2[5])	/ 100.0
    self.Canvas.itemconfig(self.power, text='Forw  %3.0fW' % power)
    frac = power / MAX_POWER
    frac = min(1.0, frac)
    self.power_meter[3] = self.power_meter[1] + self.power_meter_size * frac
    self.Canvas.coords(*self.power_meter)
    # Set Reverse power display
    refl = (self.param1[18] * 256 + self.param2[18]) / 100.0
    self.Canvas.itemconfig(self.refl, text='Refl    %3.0f' % refl)
    frac = refl / MAX_POWER
    frac = min(1.0, frac)
    self.refl_meter[3] = self.refl_meter[1] + self.refl_meter_size * frac
    self.Canvas.coords(*self.refl_meter)
    # Set SWR display
    swr = self.param2[6]	# swr code = 256 * p**2
    if power >= 2.0 and swr is not None:
      swr = math.sqrt(swr / 256.0)
      swr = (1.0 + swr) / (1.0 - swr)
      if swr > 99.9:
        swr = 99.9
      self.Canvas.itemconfig(self.swr, text='SWR  %2.1f' % swr)
      frac = (swr - 1.0) / (MAX_SWR - 1.0)
      frac = min(1.0, frac)
      self.swr_meter[3] = self.swr_meter[1] + self.swr_meter_size * frac
      self.Canvas.coords(*self.swr_meter)
    else:
      self.Canvas.itemconfig(self.swr, text='SWR')
      self.swr_meter[3] = self.swr_meter[1]
      self.Canvas.coords(*self.swr_meter)
    # Show standby/active button
    if self.is_standby:
      self.standby.Display(1)
      self.standby.config(text="Standby")
    else:
      self.standby.Display(0)
      self.standby.config(text="Active")
    # Show current antenna button
    if self.param1[4]:		# Antenna 2
      self.antenna.Display(1)
      self.antenna.config(text="Ant 2")
    else:
      self.antenna.Display(0)
      self.antenna.config(text="Ant 1")
    # Show autotune button
    if self.param1[17]:
      self.btntune.Display(0)
      self.btntune.config(text="Auto ON")
    else:
      self.btntune.Display(1)
      self.btntune.config(text="Auto OFF")
    # Show SWR threshold
    self.swrButns.DisplayIndex(self.param1[16])
    # Set status line display
    a = self.param1
    b = self.param2
    if a[3]:
      hilow = 'Low Z'
    else:
      hilow = 'High Z'
    # Freq measured period in units of 1.6usec and scaled by 32768
    # scale factor value: 32.768/1.6e-6 = 20480000
    freq_code =  a[7] * 256 + b[7]
    freq_khz = 20480000.0 / freq_code
    t = "Freq %d kHz,   Induc %d,   Capac %d,   %s,   Tune %s" % (
          freq_khz, a[1], a[2], hilow, self.tune_status)
    self.Canvas.itemconfig(self.status1, text=t)

  def About(self): # for those not inclined to RTFC (c == code) :-)
#      s = 'LDG AT-200PC Control Script\n'
#      s = s + 'Copyright (C) 2008-2010 by James C. Ahlstrom, N2ADR. All rights reserved.\n\n'
      s = '\nCopyright (C) 2008-2010 by James C. Ahlstrom, N2ADR. All rights reserved.\n\n'
      s = s + 'Copyright (C) 2019 by Christopher Sylvain, KB3CS. All rights reserved.\n\n'
      s = s + 'This free software is licensed for use under the GNU General Public License (GPL),\n'
      s = s + 'see http://opensource.org/licenses/alphabetical \n\n'
      s = s + 'Note that there is NO WARRANTY AT ALL. USE AT YOUR OWN RISK!!\n'
#      showinfo(None, s)
#      messagebox.showinfo('LDG AT-200PC Control Program', s)
      t = tkinter.Toplevel()
      t.geometry("640x240")
      t.title("LDG AT-200PC Control Program")
      msg = tkinter.Message(t, text=s, width=630, font="bold")
      msg.pack()
      b = tkinter.Button(t, text="Dismiss", command=t.destroy)
      b.pack()

if __name__ == "__main__":
  Application().main()
