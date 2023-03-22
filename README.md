replace_fonts
=============

replace_fonts replaces and unifies fonts in PowerPoint presentation.

When creating a presentation in PowerPoint, if you copy slides from a presentation created by someone else, the fonts may be different.

We want to unify fonts to create easy-to-understand presentation, but this is not easy for the following reasons.

* If you change fonts in the slide master, fonts added in text boxes, etc., are out of the scope of change.
* PowerPoint's "Replace Fonts" function requires specifying the replaced fonts one by one for all fonts used.
* Some combinations of fonts cannot be specified before and after in PowerPoint's "Replace Fonts" function.

replace_fonts solves this problem. replace_fonts fixes the fonts in a PowerPoint presentation to the fonts of the theme of that presentation (default fonts).

Requirements
------------

replace_fonts require python-pptx.

```console
pip install python-pptx
```

Usage
-----

Specify PowerPoint files as an argument.

```console
python replace_fonts.py PPT_FILE...
```

* replace_fonts will back up the specified file sequentially, open the file, replace the fonts, and save it. (`sample.pptx` is backed up to `sample - backup.pptx`)
* The replacement status is not only displayed on the screen, but also logged in a log file with the same name as the PowerPoint file. (Font replacements in `sample.pptx` will be logged in `sample.log`)
* The meanings of the theme fonts recorded in the log are as follows

  Font   | Meanings
  -------|------------------------------------------------
  +mj-lt | Heading Font Latin (Major Latin Font)
  +mn-lt | Body Font Latin (Minor Latin Font)
  +mj-ea | Heading Font East Asian (Major East Asian Font)
  +mn-ea | Body Font East Asian (Minor East Asian Font)

Tips
----

If you are using Windows, you may find it useful to add a shortcut to SendTo by doing the following

```console
pip install pywin32
python create_sendto_shortcut.py
```

Author
------

[Shinichi Akiyama](https://github.com/shakiyam)

License
-------

[MIT License](https://opensource.org/licenses/mit)
