direction = ('north', 'south', 'east', 'west', 'down', 'up', 'left', 'right', 'back')
verbs = ('go', 'stop', 'kill', 'eat')
stop = ('the', 'in', 'of', 'from', 'at', 'it')
noun = ('door', 'bear', 'princess', 'cabinet')

def scan(stuff):
    words = stuff.split()
    sentence = []

    for word in words:
        try:
            if word in direction:
                sentence.append(('direction',word))
            elif word in verbs:
                sentence.append(('verb',word))
            elif word in stop:
                sentence.append(('stop',word))
            elif word in noun:
                sentence.append(('noun',word))
            else:
                word = int(word)
                sentence.append(('number',word))
        except ValueError:
            sentence.append(('error',word))
    return sentence
