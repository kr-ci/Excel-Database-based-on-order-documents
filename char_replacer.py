def polishchar(text):
#ą, ć, ę, ł, ń, ó, ś, ź, ż.
#Ą, Ć, Ę, Ł, Ń, Ó, Ś, Ź, Ż.
#There are no problems with decoding characters "ó" and "ż", so they were not used.
    text=text.replace('¹','ą')
    text=text.replace('æ','ć')
    text=text.replace('ê','ę')
    text=text.replace('³','ł')
    text=text.replace('ñ','ń')
    text=text.replace('œ','ś')
    
    text=text.replace('¿','ż')
    text=text.replace('¥','Ą')
    text=text.replace('Æ','Ć')
    text=text.replace('Ê','Ę')
    text=text.replace('£','Ł')
    text=text.replace('Ñ','Ń')
    text=text.replace('„','Ś')
    text=text.replace('‘','Ź')
    text=text.replace('¯','Ż')

    return text
