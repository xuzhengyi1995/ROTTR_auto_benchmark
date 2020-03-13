# ROTTR_auto_benchmark

Auto run benchmark of Rise of the Tomb Raider

# Description

This script will use opencv to find the button, run the benchmark of Rise of the Tomb Raider automatically, collect the result of benchmark, finally save the result to Json file and terminate the game.

# How to use

0.  Install the following python package.

    ```shell
    pip install opencv-python
    pip install imutils
    pip install pywin32
    ```

1.  Edit the `settings.py` file.

    It's easy to understand by it's name.
    Normally only need to change the `STEAM_PATH`.

    ```python
    STEAM_PATH = '"D:\\Program Files (x86)\\Steam\\Steam.exe"'
    ROTTR_APP_ID = '391220'
    ROTTR_TITLE = 'Rise of the Tomb Raider v1.0 build 820.0_64'
    WAIT_FOR_LANUCH = 30
    ```

2.  Run `python main.py`.

# Result

The json file will like below:

```json
{
  "SpineOfTheMountain": {
    "min_fps": "28.3",
    "max_fps": "93.1",
    "avg_fps": "47.2",
    "num_frames": "1387"
  },
  "ProphetsTomb": {
    "min_fps": "28.9",
    "max_fps": "41.4",
    "avg_fps": "38.1",
    "num_frames": "935"
  },
  "GeothermalValley": {
    "min_fps": "26.8",
    "max_fps": "73.6",
    "avg_fps": "44.3",
    "num_frames": "1307"
  }
}
```
