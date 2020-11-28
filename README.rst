=============================================================
docx_text2link: Convert text in a DOCX document to hyperlinks
=============================================================

Convert specific text in a DOCX document to hyperlinks, without changing anything else.

M.H.V. Werts, May 2020, using code snippets by others cited below (many thanks!).

**USE AT YOUR OWN RISK!** This program has only been very partially tested.


-----------
Description
-----------

A typical use case for this Python script would be that you have an MS Word DOCX document with a formatted bibliography (e.g. using `Zotero`_ followed by *unlinking* the references) and you need to insert clickable hyperlinks to the corresponding papers on-line (e.g. using their Digital Object Identifiers `DOI`_). 

With Zotero it is not possible to insert such hyperlinks directly into the Word document (if you know how to do it, please tell me!). The present script provides a way to obtain hyperlinks from DOI identifiers inserted as plain text into the bibliography. Zotero can generate such plain text DOIs.

With some effort, this Python script can probably be adapted for other use cases where one would need to convert plain text in a document into hyperlinks. 

.. _Zotero: https://www.zotero.org
.. _DOI: https://www.doi.org/


-----
Usage
-----

.. code-block::

   python3 docx_text2link.py <name of input file> <name of output file>


Only one DOI per paragraph is processed. Each DOI needs to be in a separate paragraph. It may be necessary to fine-tune the script by editing it to suit your specific use case.


-------
Example
-------

One example is provided in the form of ``example_bibliography_input.docx`` which has an ("unlinked") Zotero-formatted bibliography. Using the present ``docx_text2link``, the DOIs have been converted to hyperlinks. The result of running the script is in ``example_bibliography_output.docx``.

The example input and output are also provided as PDF so that you can see the effect of the script directly by opening these documents in Github.



------------
Installation
------------

There is no specific installation procedure. The script is copied into the directory with the document to be processed, and then run by calling the ``python3`` interpreter

The script, however, relies on `python-docx`_ (we used 0.8.10), which needs to be installed first.

.. _python-docx: https://python-docx.readthedocs.io

`python-docx`_ is available on the `conda-forge`_ channel. Install with ``conda install python-docx`` . For those working with ``pip`` , there is ``pip3 install python-docx`` .

.. _conda-forge: https://conda-forge.org/



----------------
Acknowledgements
----------------


Writing of this script was possible thanks to the following information:

[1]  https://github.com/python-openxml/python-docx/issues/74

[2]  https://stackoverflow.com/questions/40475757/how-to-extract-the-url-in-hyperlinks-from-a-docx-file-using-python

[3]  https://github.com/python-openxml/python-docx/issues/74#issuecomment-261169410

The following very helpful code was used in the script:

[4]  https://github.com/python-openxml/python-docx/issues/519#issuecomment-441710870

[5]  https://github.com/python-openxml/python-docx/issues/74#issuecomment-441351994

Method [3] provides alternative way of inserting links. However, this 
generates some complications with formatting, and links can only
appear at end of a paragraph.


