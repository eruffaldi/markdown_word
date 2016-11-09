# Issue Link with text in paper
# Run with newline
# \href
#
#https://cloud.githubusercontent.com/assets/18180/20060827/4fc5a894-a4fd-11e6-984d-08511fa2c001.png
#[https://cloud.githubusercontent.com/assets/18180/20060827/4fc5a894-a4fd-11e6-984d-08511fa2c001.png](https://cloud.githubusercontent.com/assets/18180/20060827/4fc5a894-a4fd-11e6-984d-08511fa2c001.png)


import mistune
import sys

# process text if any
# emit link around text if any
class Link:
    def __init__(self,link,title,text,auto):
        self.link = link
        self.title = title
        if text is None or text == "":
            text = link
        self.text = text
        self.auto = auto

# process text and then apply style
class Style:
    def __init__(self,text,mode):
        self.text = text
        self.mode = mode

# add image needs download
# add caption
# add reference
class Image:
    def __init__(self,src,title,text):
        self.src = src
        self.title = title
        self.text = text

class Footnote:
    def __init__(self,content):
        self.text = content

class Bookmark:
    def __init__(self,label):
        self.label = label

class Run:
    def __init__(self,text):
        self.text = text

class Cite:
    def __init__(self,cite):
        self.cite = cite

class Caption:
    def __init__(self,ref):
        self.ref = ref

class Ref:
    def __init__(self,ref):
        self.ref = ref

class Holder:
    def __init__(self):
        self.runs = []
    def __iadd__(self,r):
        if type(r) is list:
            self.runs.extend(r)
        elif type(r) == Holder:
            self.runs.extend(r.runs)
        else:
            if type(r) is str:
                r = Run(r)
            self.runs.append(r)
        return self
    def __len__(self):
        return len(self.runs)

class Paragraph(Holder):
    def __init__(self,ho):
        Holder.__init__(self)
        self.runs = ho.runs
        # Analize for Images
        if len(self.runs) >= 3:
            print [x.__class__ for x in self.runs]
            # Link == the image
            # Run dummy
            # Caption 
            # rest
            if isinstance(ho.runs[0],Link) and isinstance(ho.runs[1],Run) and isinstance(ho.runs[2],Caption):
                img = Image(ho.runs[0].link,ho.runs[2].ref,ho.runs[3])
                print "replace image ignore:",ho.runs[1].text,"with img"
                self.runs = [img]

class List(Holder):
    def __init__(self,ho,ord):
        Holder.__init__(self)
        self.runs = ho.runs
        self.ordered = ord

class ListItem(Holder):
    def __init__(self,ho):
        Holder.__init__(self)
        self.runs = ho.runs

class Header:
    def __init__(self,text,level,raw=None):
        self.text = text
        self.level = level
        self.raw = raw

class Document(Holder):
    def __init__(self,r):
        Holder.__init__(self)
        self.runs = r.runs

class WordRenderer(mistune.Renderer):
    def block_code(self, code, lang):
        print "*code",len(code),lang
        return ""
    def placeholder(self):
        return Holder()
    def header(self, text, level,raw=None):
        print "*head",text,level
        return Header(text,level,raw)
    def link(self, link, title, text):
        print "*link",link,"title=",title,"text=",text
        # TODO download image and cache
        return Link(link,title,text,False)
    def autolink(self, link, is_email=False):
        print "*alink",link
        # TODO detect image and cache (if start of line)
        return Link(link,"","",True)
    def image(self, src, title, text):
        print "*image",title,src,text
        # TODO download image and cache
        return Image(src,title,text)
    def paragraph(self, text):
        # prints the RENDERED
        print "*paragraph",len(text),text
        return Paragraph(text)
    def list(self, body, ordered=True):
        # native
        return List(body,ordered)
    def list_item(self, text):
        # rendered
        return ListItem(text)
    def double_emphasis(self, text):
        return Style(text,"bold")
    def emphasis(self, text):    
        return Style(text,"italics")
    def text(self, text):
        # native
        # FIX \cite{}
        # FIX \ref{}
        # AUTO detect caption
        print "*text",text
        w = [text]
        # order matters, footnote last
        pieces = [("\\cite{",Cite),("\\ref{",Ref),("\\caption{",Caption),("\\label{",Bookmark),("\\url{",lambda x: Link(x,None,None,None)),("\\footnote",Footnote)]
        for p,t in pieces:
            r = []
            for text in w:
                if type(text) is str:
                    a = text.split(p)
                    if len(a) == 1:
                        r.append(a[0])
                    else:
                        if len(a[0]) != 0:
                            r.append(a[0])
                        for i in range(1,len(a)):
                            # text cite}text cite}text 
                            aw = a[i].split("}",1)
                            r.append(t(aw[0]))
                            if len(aw) == 2:
                                r.append(Run(aw[1]))
                else:
                    r.append(text)
            w = r
        return [Run(x) if type(x) is str else x for x in w]

#miss list paragraph 
def dump(r,w=""):
    if isinstance(r,Holder):
        print w,"holder",r
        w += " "
        for k in r.runs:
            dump(k,w)
    else:
        print w,r
def main():
    renderer = WordRenderer()
    markdown = mistune.Markdown(renderer=renderer)  
    r = Document(markdown(open(sys.argv[1],"rb").read()))
    print type(r)
    # output is holder
    print r
    dump(r)


if __name__ == '__main__':
    main()
