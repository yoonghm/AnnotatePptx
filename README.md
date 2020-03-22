# AnnotatePptx - Generate voice annotated PowerPoint Slides from Speaker Notes

This Python script generates voice annotated PowerPoint slides from embedded speaker note in each slide.

You need to have the following to use the prorgram

1.  Python 3.x
2.  Python libraries: [`google-cloud-texttospeech`](https://googleapis.dev/python/texttospeech/latest/index.html), [`python-pptx`](https://pypi.org/project/python-pptx/)
3.  A service account credential to access a project hosted in Google Cloud that with Google Cloud Text-to-Speech API enabled
4.  Set environment variable `GOOGLE_APPLICATION_CREDENTIALS` which points to a JSON file for that service account credential

You can look at [this slides deck](https://github.com/yoonghm/AnnotatePptx/blob/master/TTS.pdf) on the details.

## Usage

```bash
usage: AnnotatePptx.py [-h] [--gender [GENDER]] [--lang [LANG]]
                       [--speed [SPEED]]
                       source_pptx output_pptx

Voice annotate PowerPoint file

positional arguments:
  source_pptx        Input PowerPoint file
  output_pptx        Output PowerPoint file

optional arguments:
  -h, --help         show this help message and exit
  --gender [GENDER]  Gender (default: female)
  --lang [LANG]      Language (default: en-GB)
  --speed [SPEED]    Speed (default: 1.0)
```

## Example

If all the requirements are meet, you could try out the script with the PowerPoint file in this repo:

```bash
python AnnotatePptx.py Zero-to-Python.pptx Zero-to-Python-1.pptx --gender male --lang en-US --speed 0.95
```

## References:
1.  "Generate voice annotated PowerPoint Slides from Speaker Notes", https://medium.com/@yoonghm/generate-voice-annotated-powerpoint-slides-from-speaker-notes-82e2cd45eb27
2.  "Google.Cloud.TextToSpeech.V1", https://googleapis.github.io/google-cloud-dotnet/docs/Google.Cloud.TextToSpeech.V1/index.html
