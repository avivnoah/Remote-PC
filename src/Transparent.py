import Tkinter as tk # Python 2
root = tk.Tk()
# The image must be stored to Tk or it will be garbage collected.
label = tk.Label(root, bg='white')
root.attributes('-alpha', 0.01)
#root.overrideredirect(True)
root.geometry("1280x800")
root.lift()
root.wm_attributes("-topmost", True)
root.wm_attributes("-disabled", True)
root.wm_attributes("-transparentcolor", "white")
label.mainloop()