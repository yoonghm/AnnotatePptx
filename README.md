# AnnotatePptx - Generate voice annotated PowerPoint Slides from Speaker Notes

This Python script generates voice annotated PowerPoint slides from embedded speaker note in each slide.

You need to have the following to use the prorgram

1.  Python 3.x
2.  Python libraries: [`google-cloud-texttospeech`](https://googleapis.dev/python/texttospeech/latest/index.html), [`python-pptx`](https://pypi.org/project/python-pptx/)
3.  A service account credential to access a project hosted in Google Cloud that with Google Cloud Text-to-Speech API enabled
4.  Set environment variable `GOOGLE_APPLICATION_CREDENTIALS` which points to a JSON file for that service account credential

You can look at [this slides deck](https://github.com/yoonghm/AnnotatePptx/blob/master/TTS.pdf) on the details.
