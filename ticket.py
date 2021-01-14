import threading
import subprocess
from pynput import keyboard
threads = []


scr = '''
Add-Type -AssemblyName System.Windows.Forms
Add-Type -AssemblyName System.Drawing
$sigdir = $env:USERPROFILE+'\AppData\Roaming\Microsoft\Signatures\'
$sig = Get-Childitem –Path $sigdir -Include '*.txt' -File -Recurse -ErrorAction SilentlyContinue
$stxt = Get-Content -Path $sig.FullName
$file = $env:USERPROFILE+'\Pictures\MITE_snapshot.jpg'
if ([Windows.Forms.Clipboard]::ContainsImage()) {
  $img = [Windows.Forms.Clipboard]::GetImage()
  $img.Save($file, [Drawing.Imaging.ImageFormat]::Jpeg)
} else { 
    #return error! 
}
$Outlook = New-Object -ComObject Outlook.Application
$Mail = $Outlook.CreateItem(0)
$Mail.To = "helpdesk@euromix.in.ua"
$Mail.Subject = "проблема"
$Mail.Body = $stxt
$Mail.HTMLBody
$mail.Attachments.Add($file)
$inspector = $mail.GetInspector
$inspector.Display()
'''


def create_tictet():
    startupinfo = None
    startupinfo = subprocess.STARTUPINFO()
    startupinfo.dwFlags |= subprocess.STARTF_USESHOWWINDOW
    p = subprocess.Popen(["powershell", '-WindowStyle', 'Hidden', "-executionpolicy", "Unrestricted", scr], startupinfo=startupinfo, stdout=subprocess.PIPE)
    p = p.communicate()[0].decode("utf-8", "replace")


def on_press(key):
    try:
        if key == keyboard.Key.print_screen:
            print("Нажал принт скрин")
            create_tictet()
    except AttributeError:
        if key == keyboard.Key.print_screen:
            print("Нажал принт скрин")
            create_tictet()

def key_watch():
    with keyboard.Listener(on_press=on_press) as listener:
        listener.join()

t = threading.Thread(target=key_watch)
threads.append(t)
t.start()

