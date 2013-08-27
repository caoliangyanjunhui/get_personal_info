# This Python file uses the following encoding: utf-8
from convert_word_to_txt import WordConvert
import os, sys, codecs, simplejson

START_STR = '姓名'#.decode('UTF-8')
SEC_START_STR = '中文姓名'


class PersonalInfo(object):
	"""docstring for PersonalInfo"""
	def __init__(self, path, filename):
		super(PersonalInfo, self).__init__()
		self.__path = path
		self.__filename = filename

	def __array_to_dict(self, arr):
		info_dict = {}
		for i in range(0, len(arr), 2):
			print 'i = %d' % i
			if i == len(arr) - 1:
				break
			info_dict[arr[i]] = arr[i + 1]
		return info_dict

	def __txt_to_json(self, txt_array):
		j = 0
		info_dict = {}
		tmp = 1
		for i in range(1,len(txt_array)):
			#print '[%s, %s]' % (repr(txt_array[i]), type(txt_array[i]))
			#print '%s, %s, %s' % (type(txt_array[i]), type(START_STR.endcode('UTF-8')), type(SEC_START_STR.endcode('utf-8')))
			if txt_array[i].encode('UTF-8') == START_STR or txt_array[i].encode('UTF-8') == SEC_START_STR:
				print 'tmp;%d, i:%d' % (tmp, i)
				info_dict[j] = self.__array_to_dict(txt_array[tmp:i])
				tmp = i
				j += 1
		return simplejson.dumps(info_dict)

	def get_info(self):
		code_type = sys.getfilesystemencoding()
		print 'code type:' + code_type
		full_name = os.path.join(self.__path, self.__filename)
		txt_array = []
		file_obj = None
		try:
			file_obj = codecs.open(full_name, 'r', 'GBK')
			txt_array = file_obj.readlines()
		except Exception, e:
			print 'read txt file error: %s' % e
		finally:
			if file_obj:
				file_obj.close()

		for i in range(len(txt_array)):
			txt_array[i] = txt_array[i].encode('UTF-8').strip()
			print '<<<<txt:%s>>>>' % (txt_array[i])
			#if txt_array[len(txt_array) - 1] == '\r':
			#	txt_array = txt_array[:len(txt_array) - 1]
			if not cmp(txt_array[i], START_STR):
				print 'i = %d' % i
				txt_array = txt_array[i:]
				break
		for i in range(len(txt_array)):
			txt_array[i] = txt_array[i].strip()
			#print '<%s, %s>' % (repr(txt_array), repr(''))
			if not cmp(txt_array[i], ''):
				txt_array.remove(txt_array[i])
				continue
	
		#print txt_array
		for i in range(len(txt_array)):
			print ('<<<[%d] %s, %s>>>' % (i, txt_array[i], repr(txt_array[i])))

		return self.__txt_to_json(txt_array)
def main():
	#path = raw_input('path:')
	#filename = raw_input('file name:')
	#pinfo = PersonalInfo(path, filename)
	wdConvert = WordConvert('C:\\Users\\liacao7\\Dropbox\\Python\\0828-0907-周骅倩'.decode('UTF-8'), '入台证申请表.doc'.decode('UTF-8'))
	wdConvert.convert()
	pinfo = PersonalInfo('./0828-0907-周骅倩'.decode('UTF-8'), '入台证申请表.txt'.decode('UTF-8'))
	print pinfo.get_info()

if __name__ == '__main__':
	main()