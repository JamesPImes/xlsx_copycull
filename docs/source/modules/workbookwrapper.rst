``WorkbookWrapper``
===================

(Implemented at ``xlsx_copycull.xlsx_copycull.WorkbookWrapper`` but
automatically imported as a top-level class,
``xlsx_copycull.WorkbookWrapper``.)

.. note::

    By default, creating a ``WorkbookWrapper`` will automatically copy
    the target spreadsheet, and the original spreadsheet will not be
    modified. However, this behavior can be suppressed at initialization
    with ``no_copy=True``, in which case no copy will be generated, and
    and changes will be made to the original file instead.


.. autoclass:: xlsx_copycull.WorkbookWrapper
    :members:
    :special-members: __init__, __getitem__
