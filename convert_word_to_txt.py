import fnmatch, os, sys, simplejson, platform




class PlatformConverter(object):
    """docstring for PlatformConverter"""
    def __init__(self, path, filename):
        super(PlatformConverter, self).__init__()
        self.__path = path
        self.__filename = filename
        sysstr = platform.system()
        if sysstr == 'Windows':
            import win32com.client
        elif sysstr == 'Linux':
            pass
        elif sysstr == 'Mac':
            pass

    def save_to(self, path, filename):
        if sysstr == 'Windows':
            try:
                doc = os.path.abspath(os.path.join(self.__path, self.__filename))
                wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")
                wordapp.Documents.Open(doc)
                doc_to_txt = os.path.join(path, filename)
                wordapp.ActiveDocument.SaveAs(doc_to_txt, FileFormat = win32com.client.constants.wdFormatText)
            except Exception, e:
                raise 'win save to failed: %s' % e
            finally:
                wordapp.ActiveDocument.Close()
                wordapp.Quit()
        elif sysstr == 'Linux':
            pass
        elif sysstr == 'Mac':
            pass

class WordConvert(object):
    """docstring for WordConvert"""
    def __init__(self, path, filename):
        super(WordConvert, self).__init__()
        self.__path = path
        self.__filename = filename
        self.__pConverter =  PlatformConverter(path, filename)       

    def convert(self, target_path='', target_name = ''):
        try:
            doc = os.path.abspath(os.path.join(self.__path, self.__filename))
            self.__wordapp.Documents.Open(doc)
            ext_len = 0
            if fnmatch.fnmatch(self.__filename, '*.doc'):
                ext_len = 3
            elif fnmatch.fnmatch(self.__filename, '*.docx'):
                ext_len = 4
            if target_name == '':
                target_name = self.__filename[:-ext_len] + 'txt'
            if target_path == '':
                target_path = self.__path
            doc_to_txt = os.path.join(target_path, target_name)
            self.__pConverter.save_to(target_path, target_name)
        except Exception, e:
            raise 'convert from "%s" to "%s" failed: %s' % (self.__filename, target_name, e) 

'''class WordConvert(object):
    """docstring for WordConvert"""
    def __init__(self, path, filename):
        super(WordConvert, self).__init__()
        self.__path = path
        self.__filename = filename
        self.__wordapp = win32com.client.gencache.EnsureDispatch("Word.Application")        

    def convert(self, target_path='', target_name = ''):
        try:
            doc = os.path.abspath(os.path.join(self.__path, self.__filename))
            self.__wordapp.Documents.Open(doc)
            ext_len = 0
            if fnmatch.fnmatch(self.__filename, '*.doc'):
                ext_len = 3
            elif fnmatch.fnmatch(self.__filename, '*.docx'):
                ext_len = 4
            if target_name == '':
                target_name = self.__filename[:-ext_len] + 'txt'
            if target_path == '':
                target_path = self.__path
            doc_to_txt = os.path.join(target_path, target_name)
            self.__wordapp.ActiveDocument.SaveAs(doc_to_txt, FileFormat = win32com.client.constants.wdFormatText)
        except Exception, e:
            raise 'convert from "%s" to "%s" failed: %s' % (self.__filename, target_name, e)
        finally:
            self.__wordapp.ActiveDocument.Close()

    def __del__(self):
        self.__wordapp.Quit()
    '''
def main():
    path = raw_input("input path:")
    filename = raw_input('input filename:')
    try:
        wordConvert = WordConvert(path, filename)
        wordConvert.convert()
    except Exception, e:
        print '[Error]: %s' % e
    
if __name__ == '__main__':
    main()