import re
import shlex
from datetime import date, time, datetime
from copy import deepcopy
import warnings
from lxml.etree import Element
from lxml import etree
from zipfile import ZipFile, ZIP_DEFLATED
import shlex
import json

NAMESPACES = {
    'w': 'http://schemas.openxmlformats.org/wordprocessingml/2006/main',
    'mc': 'http://schemas.openxmlformats.org/markup-compatibility/2006',
    'ct': 'http://schemas.openxmlformats.org/package/2006/content-types',
}

CONTENT_TYPES_PARTS = (
    'application/vnd.openxmlformats-officedocument.wordprocessingml.document.main+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.header+xml',
    'application/vnd.openxmlformats-officedocument.wordprocessingml.footer+xml',
)

CONTENT_TYPE_SETTINGS = 'application/vnd.openxmlformats-officedocument.wordprocessingml.settings+xml'


class MailMerge(object):
    def __init__(self, file, remove_empty_tables=False):
        self.zip = ZipFile(file)
        self.parts = {}
        self.settings = None
        self._settings_info = None
        self.remove_empty_tables = remove_empty_tables

        try:
            content_types = etree.parse(self.zip.open('[Content_Types].xml'))
            for file in content_types.findall('{%(ct)s}Override' % NAMESPACES):
                type = file.attrib['ContentType' % NAMESPACES]
                if type in CONTENT_TYPES_PARTS:
                    zi, self.parts[zi] = self.__get_tree_of_file(file)
                elif type == CONTENT_TYPE_SETTINGS:
                    self._settings_info, self.settings = self.__get_tree_of_file(file)

            to_delete = []

            for part in self.parts.values():

                for parent in part.findall('.//{%(w)s}fldSimple/..' % NAMESPACES):
                    for idx, child in enumerate(parent):
                        if child.tag != '{%(w)s}fldSimple' % NAMESPACES:
                            continue
                        instr = child.attrib['{%(w)s}instr' % NAMESPACES]

                        args = self.__parse_field(instr)
                        name = self.__parse_instr(args)
                        print('args', args)
                        if name is None:
                            name = ''
                        parent[idx] = Element('MergeField', kind=args[0].upper(), name=name, data=json.dumps(args))

                for parent in part.findall('.//{%(w)s}instrText/../..' % NAMESPACES):
                    children = list(parent)
                    fields = zip(
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="begin"]/..' % NAMESPACES)],
                        [children.index(e) for e in
                         parent.findall('{%(w)s}r/{%(w)s}fldChar[@{%(w)s}fldCharType="end"]/..' % NAMESPACES)]
                    )

                    for idx_begin, idx_end in fields:
                        # consolidate all instrText nodes between'begin' and 'end' into a single node
                        begin = children[idx_begin]
                        instr_elements = [e for e in
                                          begin.getparent().findall('{%(w)s}r/{%(w)s}instrText' % NAMESPACES)
                                          if idx_begin < children.index(e.getparent()) < idx_end]
                        if len(instr_elements) == 0:
                            continue

                        # set the text of the first instrText element to the concatenation
                        # of all the instrText element texts
                        instr_text = ''.join([e.text for e in instr_elements])
                        instr_elements[0].text = instr_text

                        # delete all instrText elements except the first
                        for instr in instr_elements[1:]:
                            instr.getparent().remove(instr)

                        args = self.__parse_field(instr_text)
                        name = self.__parse_instr(args)
                        if name is None:
                            name = ''
                        parent[idx_begin] = Element('MergeField', kind=args[0].upper(), name=name, data=json.dumps(args))

                        # use this so we know *where* to put the replacement
                        instr_elements[0].tag = 'MergeText'
                        block = instr_elements[0].getparent()
                        # append the other tags in the w:r block too
                        parent[idx_begin].extend(list(block))

                        to_delete += [(parent, parent[i + 1])
                                      for i in range(idx_begin, idx_end)]

            for parent, child in to_delete:
                parent.remove(child)

            # Remove mail merge settings to avoid error messages when opening document in Winword
            if self.settings:
                settings_root = self.settings.getroot()
                mail_merge = settings_root.find('{%(w)s}mailMerge' % NAMESPACES)
                if mail_merge is not None:
                    settings_root.remove(mail_merge)
        except:
            self.zip.close()
            raise

    @classmethod
    def __parse_field(cls, instr):
        return shlex.split(instr, posix=False)

    @classmethod
    def __parse_instr(cls, args):
        # print('parse', args)
        if args[0] != 'MERGEFIELD':
            return None
        name = args[1]
        if name[0] == '"' and name[-1] == '"':
            name = name[1:-1]
        return name

    def __get_tree_of_file(self, file):
        fn = file.attrib['PartName' % NAMESPACES].split('/', 1)[1]
        zi = self.zip.getinfo(fn)
        return zi, etree.parse(self.zip.open(zi))

    def write(self, file):
        # Replace all remaining merge fields with empty values
        for field in self.get_merge_fields():
            self.merge(**{field: None})

        with ZipFile(file, 'w', ZIP_DEFLATED) as output:
            for zi in self.zip.filelist:
                if zi in self.parts:
                    xml = etree.tostring(self.parts[zi].getroot())
                    output.writestr(zi.filename, xml)
                elif zi == self._settings_info:
                    xml = etree.tostring(self.settings.getroot())
                    output.writestr(zi.filename, xml)
                else:
                    output.writestr(zi.filename, self.zip.read(zi))

    def get_merge_fields(self, parts=None):
        if not parts:
            parts = self.parts.values()
        fields = set()
        for part in parts:
            for mf in part.findall('.//MergeField[@kind="MERGEFIELD"]'):
                fields.add(mf.attrib['name'])
        return fields

    def merge_templates(self, replacements, separator):
        """
        Duplicate template. Creates a copy of the template, does a merge, and separates them by a new paragraph, a new break or a new section break.
        separator must be :
        - page_break : Page Break. 
        - column_break : Column Break. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - textWrapping_break : Line Break.
        - continuous_section : Continuous section break. Begins the section on the next paragraph.
        - evenPage_section : evenPage section break. section begins on the next even-numbered page, leaving the next odd page blank if necessary.
        - nextColumn_section : nextColumn section break. section begins on the following column on the page. ONLY HAVE EFFECT IF DOCUMENT HAVE COLUMNS
        - nextPage_section : nextPage section break. section begins on the following page.
        - oddPage_section : oddPage section break. section begins on the next odd-numbered page, leaving the next even page blank if necessary.
        """

        #TYPE PARAM CONTROL AND SPLIT
        valid_separators = {'page_break', 'column_break', 'textWrapping_break', 'continuous_section', 'evenPage_section', 'nextColumn_section', 'nextPage_section', 'oddPage_section'}
        if not separator in valid_separators:
            raise ValueError("Invalid separator argument")
        type, sepClass = separator.split("_")
  

        #GET ROOT - WORK WITH DOCUMENT
        for part in self.parts.values():
            root = part.getroot()
            tag = root.tag
            if tag == '{%(w)s}ftr' % NAMESPACES or tag == '{%(w)s}hdr' % NAMESPACES:
                continue
		
            if sepClass == 'section':

                #FINDING FIRST SECTION OF THE DOCUMENT
                firstSection = root.find("w:body/w:p/w:pPr/w:sectPr", namespaces=NAMESPACES)
                if firstSection == None:
                    firstSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)
			
                #MODIFY TYPE ATTRIBUTE OF FIRST SECTION FOR MERGING
                nextPageSec = deepcopy(firstSection)
                for child in nextPageSec:
                #Delete old type if exist
                    if child.tag == '{%(w)s}type' % NAMESPACES:
                        nextPageSec.remove(child)
                #Create new type (def parameter)
                newType = etree.SubElement(nextPageSec, '{%(w)s}type'  % NAMESPACES)
                newType.set('{%(w)s}val'  % NAMESPACES, type)

                #REPLACING FIRST SECTION
                secRoot = firstSection.getparent()
                secRoot.replace(firstSection, nextPageSec)

            #FINDING LAST SECTION OF THE DOCUMENT
            lastSection = root.find("w:body/w:sectPr", namespaces=NAMESPACES)

            #SAVING LAST SECTION
            mainSection = deepcopy(lastSection)
            lsecRoot = lastSection.getparent()
            lsecRoot.remove(lastSection)

            #COPY CHILDREN ELEMENTS OF BODY IN A LIST
            childrenList = root.findall('w:body/*', namespaces=NAMESPACES)

            #DELETE ALL CHILDREN OF BODY
            for child in root:
                if child.tag == '{%(w)s}body' % NAMESPACES:
                    child.clear()

            #REFILL BODY AND MERGE DOCS - ADD LAST SECTION ENCAPSULATED OR NOT
            lr = len(replacements)
            lc = len(childrenList)
            parts = []
            for i, repl in enumerate(replacements):
                for (j, n) in enumerate(childrenList):
                    element = deepcopy(n)
                    for child in root:
                        if child.tag == '{%(w)s}body' % NAMESPACES:
                            child.append(element)
                            parts.append(element)
                            if (j + 1) == lc:
                                if (i + 1) == lr:
                                    child.append(mainSection)
                                    parts.append(mainSection)
                                else:
                                    if sepClass == 'section':
                                        intSection = deepcopy(mainSection)
                                        p   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        pPr = etree.SubElement(p, '{%(w)s}pPr'  % NAMESPACES)
                                        pPr.append(intSection)
                                        parts.append(p)
                                    elif sepClass == 'break':
                                        pb   = etree.SubElement(child, '{%(w)s}p'  % NAMESPACES)
                                        r = etree.SubElement(pb, '{%(w)s}r'  % NAMESPACES)
                                        nbreak = Element('{%(w)s}br' % NAMESPACES)
                                        nbreak.attrib['{%(w)s}type' % NAMESPACES] = type
                                        r.append(nbreak)

                    self.merge(parts, **repl)

    def merge_pages(self, replacements):
         """
         Deprecated method.
         """
         warnings.warn("merge_pages has been deprecated in favour of merge_templates",
                      category=DeprecationWarning,
                      stacklevel=2)         
         self.merge_templates(replacements, "page_break")

    def merge(self, parts=None, **replacements):
        if not parts:
            parts = self.parts.values()

        for field, replacement in replacements.items():
            if isinstance(replacement, list):
                self.merge_rows(field, replacement)
            else:
                for part in parts:
                    self.__merge_field(part, field, replacement)

    @classmethod
    def eval_strftime(cls, dt, fmt):
        if dt is None:
            return ''
        elif not isinstance(dt, (datetime, date, time, )):
            return str(dt)

        def repl(m):
            res = ''
            pattern = m[0]
            while pattern:
                # Years
                if pattern[:4] in ['yyyy', 'YYYY']:
                    res += '%Y'
                    pattern = pattern[4:]
                elif pattern[:2] in ['yy', 'YY']:
                    res += '%y'
                    pattern = pattern[2:]
                elif pattern[:1] in ['y', 'Y']:
                    res += '%Y'
                    pattern = pattern[1:]

                # Months

                elif pattern[:4] == 'MMMM':
                    res += '%B'
                    pattern = pattern[4:]
                elif pattern[:2] == 'MM':
                    res += '%m'
                    pattern = pattern[2:]
                elif pattern[:1] == 'M':
                    # Hack for non-zero padded month
                    res += str(dt.month)
                    pattern = pattern[1:]

                # Days
                elif pattern[:4] == 'dddd':
                    res += '%A'
                    pattern = pattern[4:]
                elif pattern[:2] == 'dd':
                    res += '%d'
                    pattern = pattern[2:]
                elif pattern[:1] == 'd':
                    # Hack for non-zero padded month
                    res += str(dt.day)
                    pattern = pattern[1:]
                else:
                    break
            return res

        if fmt[0] == '"' and fmt[-1] == '"':
            fmt = fmt[1:-1]

        fmt = re.sub(r'[dmMyY]+', repl, fmt)
        try:
            return dt.strftime(fmt)
        except AttributeError:
            return str(dt)

    @classmethod
    def eval_star(cls, data, option):
        # This is similiar to what Word is doing. It uses
        # default formats based on locale. This should be
        # consistent behaviour.
        output = None
        if isinstance(data, (datetime, )):
            # i.e. 08/16/1988 13:42
            output = data.strftime('%x %X')
        elif isinstance(data, (date, )):
            # i.e. 08/16/1988
            output = data.strftime('%x')
        elif isinstance(data, (time, )):
            # i.e. 13:18:00
            output = data.strftime('%X')
        elif data is None:
            output = ''
        else:
            output = str(data)

        assert output is not None

        # Format
        if option.lower() == 'upper':
            output = output.upper()
        elif option.lower() == 'lower':
            output = output.lower()
        elif option.lower() == 'firstcap':
            output = output.capitalize()
        elif option.lower() == 'caps':
            output = output.title()

        return output

    @classmethod
    def eval(cls, data, params):
        print('eval', data, params)
        params = params[2:]
        evaluated = False
        for i, param in enumerate(params):
            print('param', param)
            if param == '\\@':
                data = cls.eval_strftime(data, params[i + 1])
                print('evaluated', data)
                evaluated = True
            elif param == '\\*':
                data = cls.eval_star(data, params[i + 1])
                evaluated = True

        # According to Word lack of "\* MERGEFORMAT" in MERGEFIELD is perfectly fine
        # so we treat lack of evaluation as \* code.
        if not evaluated:
            data = cls.eval_star(data, '')
        return data

    def __merge_field(self, part, field, text):
        for mf in part.findall('.//MergeField[@name="%s"]' % field):
            children = list(mf)
            # Original code is saved in "data" attribute
            instr = json.loads(mf.attrib['data'])

            mf.clear()  # clear away the attributes
            mf.tag = '{%(w)s}r' % NAMESPACES
            mf.extend(children)

            nodes = []
            # preserve new lines in replacement text
            text = self.eval(text, instr)
            text_parts = text.replace('\r', '').split('\n')
            for i, text_part in enumerate(text_parts):
                text_node = Element('{%(w)s}t' % NAMESPACES)
                text_node.text = text_part
                nodes.append(text_node)

                # if not last node add new line node
                if i < (len(text_parts) - 1):
                    nodes.append(Element('{%(w)s}br' % NAMESPACES))

            ph = mf.find('MergeText')
            if ph is not None:
                # add text nodes at the exact position where
                # MergeText was found
                index = mf.index(ph)
                for node in reversed(nodes):
                    mf.insert(index, node)
                mf.remove(ph)
            else:
                mf.extend(nodes)

    def merge_rows(self, anchor, rows):
        table, idx, template = self.__find_row_anchor(anchor)
        if table is not None:
            if len(rows) > 0:
                del table[idx]
                for i, row_data in enumerate(rows):
                    row = deepcopy(template)
                    self.merge([row], **row_data)
                    table.insert(idx + i, row)
            else:
                # if there is no data for a given table
                # we check whether table needs to be removed
                if self.remove_empty_tables:
                    parent = table.getparent()
                    parent.remove(table)

    def __find_row_anchor(self, field, parts=None):
        if not parts:
            parts = self.parts.values()
        for part in parts:
            for table in part.findall('.//{%(w)s}tbl' % NAMESPACES):
                for idx, row in enumerate(table):
                    if row.find('.//MergeField[@name="%s"]' % field) is not None:
                        return table, idx, row
        return None, None, None

    def __enter__(self):
        return self

    def __exit__(self, type, value, traceback):
        self.close()

    def close(self):
        if self.zip is not None:
            try:
                self.zip.close()
            finally:
                self.zip = None
