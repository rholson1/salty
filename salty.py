""" Compute inter-rater reliability scores by comparing SALT files
written for Python 3.4
2016-12-07 Robert H. Olson, Ph.D. rolson@waisman.wisc.edu
"""

import difflib
import re
import csv
import tkinter
from tkinter import ttk
from tkinter.filedialog import *
from tkinter import messagebox
from collections import Counter, OrderedDict
import xlsxwriter
import chardet

USE_EDITOR = True
USE_SHUFFLER = True
SIMPLE_END_PUNCT = False  # Are . and ! equivalent for purposes of comparison?
USE_FORMULAS = True  # Should Excel formulas be used for summary statistics?
TRACK_PAUSE = True  # Check for presence of a colon (except for leading colon) in utterance?


class SaltyShuffler:
    """ Interactive tool to align utterances in two files
    """
    class OPS:
        CUT = 1
        PASTE = 2
        INSERT = 3
        DELETE = 4


    def __init__(self, comparator=None):
        self.comparator = comparator

        self.edits = {}

        self.root = tkinter.Tk()
        self.root.title('Edit alignment of SALT files')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        mainframe = ttk.Frame(self.root, padding="3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)

        # Use treeview control to display info; toggle cells by clicking
        self.tree = ttk.Treeview(mainframe)
        self.tree.grid(column=0, row=0, columnspan=3, sticky=(N, W, E, S))
        ysb = ttk.Scrollbar(mainframe, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(mainframe, orient='horizontal', command=self.tree.xview)
        ysb.grid(row=0, column=3, sticky='ns')
        xsb.grid(row=1, column=0, columnspan=3, sticky='ew')
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        ttk.Button(mainframe, text='Done', command=self.done_editing).grid(column=2, row=2)
        ttk.Button(mainframe, text='Cancel', command=self.on_close).grid(column=1, row=2)
        ttk.Label(mainframe, text='Ins=Insert row, Del=Delete selected row(s), C-x=cut, C-v=paste, C-z=undo, m=mark').grid(column=0, row=2)

        self.columns = ('a', 'b')
        col_labels = (self.comparator.file1, self.comparator.file2)

        self.tree['columns'] = self.columns
        for col, label in zip(self.columns, col_labels):
            self.tree.heading(col, text=label)

        # Set column widths
        self.tree.column('#0', width=40)
        for col in self.columns[2:]:
            self.tree.column(col, width=80, anchor='center')

        for idx, row in enumerate(self.comparator.report):
            if idx % 2 == 0:
                tags = ('clickable', 'even')
            else:
                tags = ('clickable',)

            self.tree.insert('', 'end', iid=str(idx), text=str(idx), values=[row[c] for c in self.columns], tags=tags)

        self.tree.tag_bind('clickable', '<ButtonPress>', self.click_handler)
        self.tree.tag_configure('even', background='#f80f80f80')
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)
        self.tree.bind('<Delete>', self.delete_selection)
        self.tree.bind('<Insert>', self.insert_row)
        self.tree.bind('<Control-x>', self.cut_cell)
        self.tree.bind('<Control-v>', self.paste_cell)
        self.tree.bind('<Control-z>', self.undo)
        self.tree.bind('m', self.mark_row)

        self.current_row = None
        self.current_col = None
        self.current_item = None
        self.clipboard = ''
        self.clipboard_col = None
        self.old_clipboard = None  # for Undo
        self.last_selection = None
        self.last_op = None
        self.last_row = None
        self.deleted_items = None
        self.replaced_values = None

        self.root.mainloop()

    def undo(self, event):
        if self.last_op == self.OPS.CUT:
            self.tree.selection_set(self.last_selection)
            for item, value in zip(self.tree.selection(), self.clipboard):
                self.tree.set(item, self.clipboard_col, value)
        elif self.last_op == self.OPS.PASTE:
            # Put items back in clipboard and clear cells where they were pasted (does not restore previous contents)
            self.clipboard = self.old_clipboard
            for item, value in zip(self.tree.get_children()[self.last_row: self.last_row + len(self.clipboard)], self.replaced_values):
                self.tree.set(item, self.clipboard_col, value)
        elif self.last_op == self.OPS.DELETE:
            # Restore removed items...
            for item in reversed(self.deleted_items):
                self.tree.move(item, '', self.last_row)
            self.deleted_items = None
            self.tree.selection_set(self.last_selection)
        elif self.last_op == self.OPS.INSERT:
            # Delete inserted row
            inserted_row = self.tree.get_children()[self.last_row]
            self.tree.delete(inserted_row)
        self.last_op = None

    def cut_cell(self, event):
        if self.tree.selection():
            self.last_op = self.OPS.CUT
            self.clipboard = [self.tree.set(item, self.current_col) for item in self.tree.selection()]
            self.clipboard_col = self.current_col
            for item in self.tree.selection():
                self.tree.set(item, self.current_col, '')
            self.last_selection = self.tree.selection()

    def paste_cell(self, event):
        if self.clipboard:
            self.last_op = self.OPS.PASTE
            dest_items = self.tree.get_children()[self.current_row:self.current_row + len(self.clipboard)]
            self.replaced_values = []
            for item, value in zip(dest_items, self.clipboard):
                self.replaced_values.append(self.tree.set(item, self.clipboard_col))  # store overwritten value
                self.tree.set(item, self.clipboard_col, value)
            self.old_clipboard = self.clipboard
            self.clipboard = ''
            self.last_row = self.current_row

    def delete_selection(self, event):
        if self.tree.selection():
            self.last_op = self.OPS.DELETE
            self.deleted_items = self.tree.selection()
            self.last_row = self.tree.index(self.deleted_items[0])
            self.last_selection = self.tree.selection()
            self.tree.detach(*self.tree.selection())

    def insert_row(self, event):
        if self.tree.selection():
            self.last_op = self.OPS.INSERT
            self.current_row = self.tree.index(self.tree.selection()[0])
            self.tree.insert('', self.current_row, tags=('clickable',))
            self.last_row = self.current_row

    def mark_row(self, event):
        for item in self.tree.selection():
            if self.is_marked(item):
                self.tree.item(item, text=item)
            else:
                self.tree.item(item, text=item + " *")

    def is_marked(self, item):
        return self.tree.item(item, option='text')[-1:] == '*'

    def click_handler(self, event):
        # Record the item/row/column that was clicked on (for use in cut/paste)
        self.current_item = self.tree.identify_row(event.y)
        self.current_row = self.tree.index(self.current_item)
        self.current_col = self.tree.identify_column(event.x)

    def done_editing(self):
        report = [report_entry(self.tree.set(i, 'a'), self.tree.set(i, 'b'), self.is_marked(i)) for i in self.tree.get_children()]
        self.comparator.report = report
        self.on_close()

    def on_close(self):
        self.root.destroy()
        self.root.quit()


class SaltyEditor:
    def __init__(self, comparator=None):
        self.comparator = comparator

        self.edits = {}

        self.root = tkinter.Tk()
        self.root.title('Edit SALT comparison')
        self.root.rowconfigure(0, weight=1)
        self.root.columnconfigure(0, weight=1)
        mainframe = ttk.Frame(self.root, padding="3 3 12 12")
        mainframe.grid(column=0, row=0, sticky=(N, W, E, S))
        mainframe.columnconfigure(0, weight=1)
        mainframe.rowconfigure(0, weight=1)

        # Use treeview control to display info; toggle cells by clicking
        self.tree = ttk.Treeview(mainframe)
        self.tree.grid(column=0, row=0, columnspan=3, sticky=(N, W, E, S))
        ysb = ttk.Scrollbar(mainframe, orient='vertical', command=self.tree.yview)
        xsb = ttk.Scrollbar(mainframe, orient='horizontal', command=self.tree.xview)
        ysb.grid(row=0, column=3, sticky='ns')
        xsb.grid(row=1, column=0, columnspan=3, sticky='ew')
        self.tree.configure(yscroll=ysb.set, xscroll=xsb.set)

        ttk.Button(mainframe, text='Done', command=self.done_editing).grid(column=2, row=2)
        ttk.Button(mainframe, text='Cancel', command=self.on_close).grid(column=1, row=2)
        ttk.Label(mainframe, text='Click cell to cycle value (1, 0, blank).  Right click to reset to original value.').grid(column=0, row=2)

        self.columns = ['a', 'b', 'utt_seg', 'XX', '^or>', '()', '<>', '# morphs', '# words', 'word ID', 'end punct']
        col_labels = [self.comparator.file1, self.comparator.file2,
                      'Utt Segm', 'XX', '^ or >', '()', '<>', '# morphs', '# words', 'word ID', 'end punct']
        if TRACK_PAUSE:
            self.columns.insert(7, ':')
            col_labels.insert(7, ':')

        self.tree['columns'] = self.columns
        for col, label in zip(self.columns, col_labels):
            self.tree.heading(col, text=label)

        # Set column widths
        self.tree.column('#0', width=40)
        for col in self.columns[2:]:
            self.tree.column(col, width=80, anchor='center')

        for idx, row in enumerate(self.comparator.report):
            if idx % 2 == 0:
                tags = ('clickable', 'even')
            else:
                tags = ('clickable',)

            text = str(idx)
            if row['mark']:
                text += ' *'
                tags = ('clickable', 'highlight')
            self.tree.insert('', 'end', iid=str(idx), text=text, values=[row[c] for c in self.columns], tags=tags)

        self.tree.tag_bind('clickable', '<ButtonPress>', self.click_handler)
        self.tree.tag_bind('clickable', '<Button-3>', self.rightclick_handler)
        self.tree.tag_configure('even', background='#f80f80f80')
        self.tree.tag_configure('highlight', background='#f4f442')
        self.root.protocol("WM_DELETE_WINDOW", self.on_close)

        self.root.mainloop()

    def rightclick_handler(self, event):
        # Reset to the original value
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)
        col_num = int(column[1:]) - 1
        col_name = self.columns[col_num]
        if column not in ('#0', '#1', '#2'):
            original_value = self.comparator.report[int(row)][col_name]
            self.tree.set(row, column, str(original_value))

    def click_handler(self, event):
        column = self.tree.identify_column(event.x)
        row = self.tree.identify_row(event.y)

        col_num = int(column[1:]) - 1
        col_name = self.columns[col_num]
        if column not in ('#0', '#1', '#2'):
            value = self.tree.set(row, column)
            new_value = self.toggle(value)
            self.tree.set(row, column, new_value)
            # Store the updated value for later application
            self.edits[(int(row), col_name)] = int(new_value) if new_value else new_value

    def toggle(self, value):
        # cycle between ('1', '0', '')
        cycle = {'1': '0', '0': '', '': '1'}
        return cycle[value]

    def apply_edits(self):
        for row, column in self.edits.keys():
            self.comparator.report[row][column] = self.edits[(row, column)]

    def done_editing(self):
        self.apply_edits()
        self.comparator.compute_stats()
        #self.comparator.write_output()
        self.comparator.write_xlsx()

        self.on_close()

    def on_close(self):
        self.root.destroy()
        self.root.quit()


def compare_salt_files():
    # Prompt user for input/output files, and then run the SaltComparator

    # Suppress tkinter root window
    root = tkinter.Tk()
    root.withdraw()

    salt_files = askopenfilenames(title='Select two SALT files',
                                  filetypes=[('SALT files', '.SLT'), ('All Files','.*')])
    salt_files = root.tk.splitlist(salt_files)
    if len(salt_files) == 2:
        output_file = asksaveasfilename(title='Select output file',
                                      filetypes=[('Excel spreadsheet', '.xlsx'), ('Tab-delimited text file', '.txt'), ('All Files', '.*')])
        if output_file:
            SaltComparator(salt_files[0], salt_files[1], output_file)
    elif len(salt_files) > 0:
        messagebox.showerror('SALT Comparator', 'Two SALT files must be selected for comparison.')


def test_comparator():
    files = ('SHv07SALT_SE.SLT', 'SHv07_intraraterrely_SE.SLT')
    output = 'salt_reliability.txt'

    sc = SaltComparator(files[0], files[1], output)


def report_entry(a='', b='', mark=False):
    return {'a': a,
            'b': b,
            'utt_seg': 1 if a or b else '',  # assume correct segmentation, but check for errors
            'XX': '',
            '^or>': '',
            '()': '',
            '<>': '',
            ':': '',
            '# morphs': '',
            '# words': '',
            'word ID': '',
            'end punct': '',
            'mark': mark}


class SaltComparator:
    def __init__(self, file1, file2, output):
        self.file1 = file1
        self.file2 = file2
        self.output = output

        # Read input files
        lines = []
        for file in (file1, file2):
            # first detect file encoding using chardet, then read the file using the detected encoding
            with open(file, 'rb') as f:
                encoding = chardet.detect(f.read())['encoding']
            with open(file, 'r', encoding=encoding) as f:
                lines.append([get_speaker(L) + L[1:] for L in f.readlines() if L[0] not in ('-', '=', ':', ';')])

        lowerlines = [[s.lower() for s in lineset] for lineset in lines]

        # Use HtmlDiff to align corresponding lines
        hd = difflib.HtmlDiff()
        table = hd.make_table(lowerlines[0], lowerlines[1], file1, file2)

        # Store the matched lines in list of dictionaries
        self.report = []
        for row in table.split('\n'):
            if row.lstrip()[:4] == '<tr>':
                row_a, row_b = parse_html_row(row)  # see which rows were matched up
                line_a = lines[0][row_a - 1].strip() if row_a else ''
                line_b = lines[1][row_b - 1].strip() if row_b else ''
                line_a = "'" + line_a if line_a[:1] == '+' else line_a
                line_b = "'" + line_b if line_b[:1] == '+' else line_b
                self.report += [report_entry(line_a, line_b)]

        # Identify speakers (e.g. Child, Parent, Sibling, Other)
        speakers = get_speakers(lines[0])
        self.speaker_initials = [get_speaker(s) for s in speakers if s]
        self.stats = OrderedDict()

        if USE_SHUFFLER:
            SaltyShuffler(self)

        self.compute_segmentation()
        self.score_lines()

        if USE_EDITOR:
            SaltyEditor(self)
        else:
            self.compute_stats()
            #self.write_output()
            self.write_xlsx()


    def compute_segmentation(self):
        # Determine if each line has been segmented correctly.
        # Focus on cases where one of the lines is blank
        # Compare the non-blank line to lines before and after, accounting for speaker
        # Check unmatched lines to see if there is an adjacent segmentation error
        def clean(s):
            # prepare a string to have a ratio computed for comparison
            # 1. lower case
            # 2. single spaces separating elements
            # 3. omit trailing punctuation (last character in string)
            return ' '.join(s.lower()[:-1].split())

        def is_utterance(column, line):
            # Is the report entry associated with a speaker?
            return get_speaker(self.report[line][column]) in self.speaker_initials

        sm = difflib.SequenceMatcher()
        for idx in range(len(self.report)):

            # Clear utt_seg for non-speaker lines
            if not is_utterance('a', idx):
                self.report[idx]['utt_seg'] = ''

            # Look at unmatched lines
            if self.report[idx]['a'] == '' or self.report[idx]['b'] == '':
                if self.report[idx]['a'] == '' and is_utterance('b', idx):
                    a, b = 'b', 'a'
                elif self.report[idx]['b'] == '' and is_utterance('a', idx):
                    a, b = 'a', 'b'
                else:
                    # This line has no utterances; no action required
                    continue

                self.report[idx]['utt_seg'] = 0  # unmatched segment is always a segmentation error

                # a is the non-blank side
                # compare prev b to prev a + a
                # compare next b to a + next a
                # when combining lines, omit interior punctuation and prefix

                # Does adding the lone line to prev/next line improve ratio?  If so, assume it's a match (error)

                # Before
                try:
                    if self.report[idx][a][0] == self.report[idx - 1][a][0]:  # same speaker
                        sm.set_seq2(clean(self.report[idx - 1][b]))
                        sm.set_seq1(clean(self.report[idx - 1][a]))
                        ratio_before = sm.ratio()
                        sm.set_seq1(clean(self.report[idx - 1][a][:-1] + self.report[idx][a][1:]))
                        ratio_after = sm.ratio()
                        if ratio_after > ratio_before:
                            self.report[idx - 1]['utt_seg'] = ''
                except IndexError:
                    pass

                # After
                try:
                    if self.report[idx][a][0] == self.report[idx + 1][a][0]:  # same speaker
                        sm.set_seq2(clean(self.report[idx + 1][b]))
                        sm.set_seq1(clean(self.report[idx + 1][a]))
                        ratio_before = sm.ratio()
                        sm.set_seq1(clean(self.report[idx][a][:-1] + self.report[idx + 1][a][1:]))
                        ratio_after = sm.ratio()
                        if ratio_after > ratio_before:
                            self.report[idx + 1]['utt_seg'] = ''
                except IndexError:
                    pass

    def score_lines(self):
        # compute ratings for each utterance line
        def rate(func, a, b):
            if func(a) and func(b):
                val = 1
            elif func(a) or func(b):
                val = 0
            else:
                val = ''
            return val

        def has_xx(s):
            w = re.sub('{.+}|[<>()^.!?,:]', '', s).lower().split()
            return 'x' in w or 'xx' in w or 'xxx' in w

        def interrupted(s):
            return s[-1:] in ('^', '>')

        def has_maze(s):
            return '(' in s and ')' in s

        def has_overlap(s):
            return '<' in s and '>' in s

        def same_val(func, a, b):
            if func(a) == func(b):
                val = 1
            else:
                val = 0
            return val

        def num_morphs(s):
            # first strip out punctuation and exclude text between () or {} and words that start with *
            bare = re.sub('\([^\)]+\)|{.+}|[<>()^.!?,:"]', '', s)
            return len(bare.replace('/', ' ').split())

        def words(s):
            # first strip out punctuation and exclude text between () or {} and words that start with *
            bare = re.sub('\([^\)]+\)|{.+}|[<>()^.!?,:"]', '', s).lower()
            # exclude anything after a / within a word
            bare = re.sub('/[^ ]+', '', bare)
            return bare.split()

        def word_set(s):
            # Sort words because word order doesn't matter.
            return sorted(words(s))

        def num_words(s):
            return len(words(s))

        def end_punct(s):
            # Return the punctuation
            if SIMPLE_END_PUNCT:
                # make '!' and '.' equivalent for purposes of comparison
                return s[-1:].replace('!', '.')
            else:
                return s[-1:]

        def has_pause(s):
            # True if a colon ':' is in s anywhere after the first two positions (i.e., ignore leading colon)
            return ':' in s[2:]

        def clean(s):
            """ Remove text that should be ignored:
            1. text enclosed in brackets []
            2. text that begins with * (and optionally, leading /)
            
            Also convert smart quotes to dumb quotes
            """
            charmap = {0x201c: '"',
                       0x201d: '"',
                       0x2018: "'",
                       0x2019: "'"}

            return re.sub("/?\*['\w]+| *\[[^\]]+\]", '', s).translate(charmap)

        for idx in range(len(self.report)):
            a = clean(self.report[idx]['a'])
            b = clean(self.report[idx]['b'])
            if get_speaker(a) in self.speaker_initials and self.report[idx]['utt_seg'] == 1:

                self.report[idx]['XX'] = rate(has_xx, a, b)
                self.report[idx]['^or>'] = rate(interrupted, a, b)
                self.report[idx]['()'] = rate(has_maze, a, b)
                self.report[idx]['<>'] = rate(has_overlap, a, b)
                self.report[idx][':'] = rate(has_pause, a, b)

                if self.report[idx]['XX'] == '' and self.report[idx]['^or>'] == '':
                    self.report[idx]['# morphs'] = same_val(num_morphs, a, b)
                    self.report[idx]['# words'] = same_val(num_words, a, b)
                    if self.report[idx]['# words'] == 1:
                        self.report[idx]['word ID'] = same_val(word_set, a, b)

                self.report[idx]['end punct'] = same_val(end_punct, a, b)

    def compute_stats(self):
        # initialize data structure
        for speaker in self.speaker_initials + ['all']:
            self.stats[speaker] = Counter({x: 0 for x in ('total', 'total2', 'utt_seg', 'XX', '^or>', '()', '<>', ':',
                                                          '# morphs', '# words', 'word ID', 'end punct')})
        # gather data
        for idx in range(len(self.report)):
            row = self.report[idx]
            speaker = get_speaker(row['a']) or get_speaker(row['b'])

            self.report[idx]['speaker'] = speaker  # for formulas

            if speaker in self.speaker_initials:
                if row['utt_seg'] in (0, 1):
                    self.stats[speaker]['total'] += 1
                if row['# morphs'] in (0, 1):
                    self.stats[speaker]['total2'] += 1
                if row['utt_seg'] == 1:
                    self.stats[speaker]['utt_seg'] += 1
                    for k in ('XX', '^or>', '()', '<>', 'end punct', ':'):
                        if row[k] != 0:
                            self.stats[speaker][k] += 1
                    for k in ('# morphs', '# words', 'word ID'):
                        if row[k] == 1:
                            self.stats[speaker][k] += 1

        # assemble results for all speakers
        for speaker in self.speaker_initials:
            self.stats['all'].update(self.stats[speaker])

    def write_output(self):
        if TRACK_PAUSE:
            headers = ['Utt Segm', 'XX', '^ or >', '()', '<>', ':', '# morphs', '# words', 'word ID', 'end punct']
            colkeys = ['utt_seg', 'XX', '^or>', '()', '<>', ':', '# morphs', '# words', 'word ID', 'end punct']
        else:
            headers = ['Utt Segm', 'XX', '^ or >', '()', '<>', '# morphs', '# words', 'word ID', 'end punct']
            colkeys = ['utt_seg', 'XX', '^or>', '()', '<>', '# morphs', '# words', 'word ID', 'end punct']

        with open(self.output, 'w', newline='') as f:
            writer = csv.writer(f, dialect='excel-tab')
            writer.writerow([self.file1, self.file2] + headers)
            for row in self.report:
                if row['a'] or row['b']:
                    writer.writerow([row[k] for k in ['a', 'b'] + colkeys])

            for speaker, row in self.stats.items():
                row = self.stats[speaker]
                to_write = [
                    'Speaker summary',
                    speaker,
                    '="{} / {}"'.format(row['utt_seg'], row['total']),
                    '="{} / {}"'.format(row['XX'], row['utt_seg']),
                    '="{} / {}"'.format(row['^or>'], row['utt_seg']),
                    '="{} / {}"'.format(row['()'], row['utt_seg']),
                    '="{} / {}"'.format(row['<>'], row['utt_seg']),
                    '="{} / {}"'.format(row['# morphs'], row['total2']),
                    '="{} / {}"'.format(row['# words'], row['total2']),
                    '="{} / {}"'.format(row['word ID'], row['# words']),
                    '="{} / {}"'.format(row['end punct'], row['utt_seg'])
                ]
                if TRACK_PAUSE:
                    to_write.insert(7, '="{} / {}"'.format(row[':'], row['utt_seg']))

                writer.writerow(to_write)

    def write_xlsx(self):
        # Create xlsx file to support dynamic formulas

        if TRACK_PAUSE:
            colkeys = ['utt_seg', 'XX', '^or>', '()', '<>', ':', '# morphs', '# words', 'word ID', 'end punct']
        else:
            colkeys = ['utt_seg', 'XX', '^or>', '()', '<>', '# morphs', '# words', 'word ID', 'end punct']

        def xl_write_line(sheet, row, line, highlights=None, format=None):
            """ write an array to a spreadsheet, highlighting specified cells
            sheet : the destination worksheet
            row : 0-based row number
            line : iterable which yields values to write
            highlights : iterable which yields column numbers of cells in the row to be highlighted
            """
            if highlights is None:
                highlights = []
            for col, value in enumerate(line):
                if col in highlights:
                    sheet.write(row, col, value, format)
                else:
                    sheet.write(row, col, value)

        def xl_formula(format_string, row, speakers, *args):
            # given a format string, expand to create Excel-compatible formula
            return '(' + ' + '.join([format_string.format(row, spk, *args) for spk in speakers]) + ')'

        r = 0
        with xlsxwriter.Workbook(self.output) as workbook:
            worksheet = workbook.add_worksheet()
            headers = [self.file1, self.file2, 'Speaker',
                'Utt Segm', 'XX', '^ or >', '()', '<>', '# morphs', '# words', 'word ID', 'end punct'
            ]
            if TRACK_PAUSE:
                headers.insert(8, ':')
            xl_write_line(worksheet, r, headers)
            r += 1

            for row in self.report:
                if row['a'] or row['b']:
                    xl_write_line(worksheet, r,
                                  [row[k] for k in ['a', 'b', 'speaker'] + colkeys])
                    r += 1

            r_end = r

            col_offset = 1 if TRACK_PAUSE else 0  # Need to shift columns if inserting pause column
            for speaker, row in self.stats.items():
                if USE_FORMULAS:
                    if speaker == 'all':
                        speakers = self.speaker_initials
                    else:
                        speakers = [speaker]

                    row = {
                        'utt_seg': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1)', r_end, speakers),
                        'total': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, ">-1")', r_end, speakers),
                        'total2': xl_formula('COUNTIFS(C2:C{0}, "{1}", {2}2:{2}{0}, ">-1")', r_end, speakers, chr(ord('I') + col_offset)),
                        'XX': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, E2:E{0}, "<>0")', r_end, speakers),
                        '^or>': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, F2:F{0}, "<>0")', r_end, speakers),
                        '()': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, G2:G{0}, "<>0")', r_end, speakers),
                        '<>': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, H2:H{0}, "<>0")', r_end, speakers),
                        ':': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, I2:I{0}, "<>0")', r_end, speakers),
                        '# morphs': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, {2}2:{2}{0}, 1)', r_end, speakers, chr(ord('I') + col_offset)),
                        '# words': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, {2}2:{2}{0}, 1)', r_end, speakers, chr(ord('J') + col_offset)),
                        'word ID': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, {2}2:{2}{0}, 1)', r_end, speakers, chr(ord('K') + col_offset)),
                        'end punct': xl_formula('COUNTIFS(C2:C{0}, "{1}", D2:D{0}, 1, {2}2:{2}{0}, "<>0")', r_end, speakers, chr(ord('L') + col_offset))
                    }

                    # Summary Stats: Fraction Format
                    line = [
                        'Speaker summary',
                        speaker,
                        '',
                        '={} & " / " & {}'.format(row['utt_seg'], row['total']),
                        '={} & " / " & {}'.format(row['XX'], row['utt_seg']),
                        '={} & " / " & {}'.format(row['^or>'], row['utt_seg']),
                        '={} & " / " & {}'.format(row['()'], row['utt_seg']),
                        '={} & " / " & {}'.format(row['<>'], row['utt_seg']),
                        '={} & " / " & {}'.format(row['# morphs'], row['total2']),
                        '={} & " / " & {}'.format(row['# words'], row['total2']),
                        '={} & " / " & {}'.format(row['word ID'], row['# words']),
                        '={} & " / " & {}'.format(row['end punct'], row['utt_seg'])
                    ]
                    if TRACK_PAUSE:
                        line.insert(8, '={} & " / " & {}'.format(row[':'], row['utt_seg']))
                    xl_write_line(worksheet, r, line)
                    r += 1

                    # Summary Stats: Decimal Format
                    line = [
                        'Speaker summary',
                        speaker,
                        '',
                        '={} / {}'.format(row['utt_seg'], row['total']),
                        '={} / {}'.format(row['XX'], row['utt_seg']),
                        '={} / {}'.format(row['^or>'], row['utt_seg']),
                        '={} / {}'.format(row['()'], row['utt_seg']),
                        '={} / {}'.format(row['<>'], row['utt_seg']),
                        '={} / {}'.format(row['# morphs'], row['total2']),
                        '={} / {}'.format(row['# words'], row['total2']),
                        '={} / {}'.format(row['word ID'], row['# words']),
                        '={} / {}'.format(row['end punct'], row['utt_seg'])
                    ]
                    if TRACK_PAUSE:
                        line.insert(8, '={} / {}'.format(row[':'], row['utt_seg']))
                    xl_write_line(worksheet, r, line)
                    r += 1
                else:
                    # Static summary statistics
                    row = self.stats[speaker]
                    line = [
                        'Speaker summary',
                        speaker,
                        '',
                        '{} / {}'.format(row['utt_seg'], row['total']),
                        '{} / {}'.format(row['XX'], row['utt_seg']),
                        '{} / {}'.format(row['^or>'], row['utt_seg']),
                        '{} / {}'.format(row['()'], row['utt_seg']),
                        '{} / {}'.format(row['<>'], row['utt_seg']),
                        '{} / {}'.format(row['# morphs'], row['total2']),
                        '{} / {}'.format(row['# words'], row['total2']),
                        '{} / {}'.format(row['word ID'], row['# words']),
                        '{} / {}'.format(row['end punct'], row['utt_seg'])
                    ]
                    if TRACK_PAUSE:
                        line.insert(8, '{} / {}'.format(row[':'], row['utt_seg']))
                    xl_write_line(worksheet, r, line)
                    r += 1


def get_speakers(lines):
    """ given a line of the form '$ [speaker1], [speaker2], ...' return the list of speakers"""
    speaker_line = next((L for L in lines if L[0] == '$'))
    parts = speaker_line[1:].split(',')
    speakers = [s.strip() for s in parts]
    return speakers


def get_speaker(s):
    """ Given a string, map the first character to a speaker """
    firstchar = s[:1]
    return 'P' if firstchar in ('M', 'D') else firstchar


def parse_html_row(line):
    """ Extract information from a row of an html table"""

    # Assumes six columns
    parts = re.search('<td.*?>(.*?)</td><td.*?>(.*?)</td><td.*?>(.*?)</td><td.*?>(.*?)</td><td.*?>(.*?)</td><td.*?>(.*?)</td>', line).groups()

    # or, making no assumptions about number of columns
    # line = re.sub('</?tr>', '', line)
    # parts = re.split('</?td.+?>', line)

    try:
        line1 = int(parts[1]) if parts[1] else None
        line2 = int(parts[4]) if parts[4] else None
    except ValueError:
        line1 = line2 = None

    return line1, line2


if __name__ == '__main__':
    compare_salt_files()
