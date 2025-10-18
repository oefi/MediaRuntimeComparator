# MediaRuntimeComparator

**MediaRuntimeComparator** is a Windows PowerShell GUI tool that helps you identify duplicate or near‑duplicate media files by comparing their **runtime, size, and name**.  
It leverages `ffprobe` from the [FFmpeg](https://ffmpeg.org/) project to extract accurate duration metadata.

---

## ✨ Features
- 🖥️ **GUI interface** built with Windows Forms (no command‑line needed)
- 🔍 **Recursive scan** of folders for video and audio files
- ⏱️ Displays **duration (HH:MM:SS)**, **(milli)seconds**, and **file size**
- 🎯 Highlights potential duplicates within a configurable **tolerance (0–60s)**
- 🗑️ Right‑click to **delete files** directly from the list
- 📂 Double‑click to **open files** in your default media player
- 💾 Saves settings (last folder, window size, tolerance, ffprobe path) between sessions

---

## 📦 Requirements
- Windows with **PowerShell 5.1+** (or PowerShell 7 with Windows Forms support)
- `ffprobe.exe` (part of the FFmpeg tools, available at [ffmpeg.org/download](https://ffmpeg.org/download.html))

---

## 🚀 Installation & Usage

1. You have two options to get started:

      **Option A – Clone the repository**

      - Run `git clone https://github.com/yourusername/MediaRuntimeComparator.git` and then `cd MediaRuntimeComparator`.

      **Option B – Copy the script directly**
     
      - Navigate to the [Files](../../) section of this repository
      - open `MediaRuntimeComparator.ps1`
      - click **Raw** → then copy and paste the contents into a new file named `MediaRuntimeComparator.ps1` on your computer.
   
2. **Ensure `ffprobe.exe` is available:**  
   - Download it as part of the official [FFmpeg builds](https://ffmpeg.org/download.html).  
     On Windows, you can grab a pre‑compiled zip from [gyan.dev](https://www.gyan.dev/ffmpeg/builds/) or [BtbN builds](https://github.com/BtbN/FFmpeg-Builds/releases).  
     Inside the archive you’ll find `bin\ffprobe.exe`.  
   - Place `ffprobe.exe` in the same folder as the script, **or**  
   - Set its path in the GUI once the app is running.

3. **Run the script in your powershell terminal:**  
    - Execute `powershell -ExecutionPolicy Bypass -File .\MediaRuntimeComparator.ps1`.

4. **In the GUI:**  
   - Select a folder to scan  
   - Click **Scan**  
   - Review results, delete duplicates, or open files directly

---

## ⚡ Quick Launch (Windows 10/11 Shortcuts)

Instead of typing the PowerShell command every time, you can create a shortcut to run the script with a double‑click:

1. Right‑click on your desktop → **New** → **Shortcut**  
2. In the location field, paste the following (adjust the path to where you saved the script):  
   `powershell.exe -ExecutionPolicy Bypass -File "C:\Path\To\MediaRuntimeComparator.ps1"`  
3. Click **Next**, give it a name (e.g. `MediaRuntimeComparator`), and finish.  
4. (Optional) Right‑click the shortcut → **Properties** → **Change Icon…** to give it a custom look.  
5. Now you can launch the tool anytime with a double‑click.
<img width="812" height="527" alt="image" src="https://github.com/user-attachments/assets/1657a188-1314-4f52-9af3-e607ef89b314" />



💡 Tip: You can also **pin the shortcut to Start or Taskbar** for even faster access.

---

## ⚙️ Configuration
A config file `MediaRuntimeComparator.cfg` is automatically created in the script directory.  
It stores:
- Last used folder  
- Window size  
- ffprobe path  
- Tolerance setting  

---

## 📂 Supported Formats
The tool recognizes most common video and audio formats, including:  
`.mp4, .mkv, .avi, .mov, .wmv, .flv, .webm, .ts, .m2ts, .m4v, .mpg, .mpeg, .3gp, .3g2, .ogg, .ogv, .ogm, .vob, .divx, .rm, .rmvb, .asf, .f4v, .mxf, .mts, .m2v, .mp2, .mp3, .aac, .wav, .flac, .alac, .wma, .m4a, .opus, .aiff, .au, .ac3, .dts, .amr, .caf`

---

## 🖼️ Screenshot (GIF)

![mrcfffprobe](https://github.com/user-attachments/assets/ecf5f33d-7ed9-4826-98bc-1bf10730581f)

---

## 🤝 Contributing
Contributions, issues, and feature requests are welcome!  
Feel free to open an [issue](../../issues) or submit a pull request.

---

## ⚖️ License
This project is licensed under the **MIT License** – see the [LICENSE](LICENSE) file for details.
