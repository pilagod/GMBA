__author__ = 'pilagod'

class EnglishExam:

    def __init__(self, total=0, listening=0, speaking=0, reading=0, writing=0):
        self.__total = float(total)
        self.__listening = float(listening)
        self.__speaking = float(speaking)
        self.__reading = float(reading)
        self.__writing = float(writing)

    @property
    def total(self):
        return self.__total
    @total.setter
    def total(self, total):
        self.__total = total

    @property
    def listening(self):
        return self.__listening
    @listening.setter
    def listening(self, listening):
        self.__listening = listening

    @property
    def speaking(self):
        return self.__speaking
    @speaking.setter
    def speaking(self, speaking):
        self.__speaking = speaking

    @property
    def reading(self):
        return self.__reading
    @reading.setter
    def reading(self, reading):
        self.__reading = reading

    @property
    def writing(self):
        return self.__writing
    @writing.setter
    def writing(self, writing):
        self.__writing

    # Input:
    #   'scores':refer to a string of serial scores in format: T/L/S/R/W
    #
    def setScores(self, scores):
        scores = scores.split('/')
        if len(scores) == 5:
            self.__init__(float(scores[0]), float(scores[1]), float(scores[2]), float(scores[3]), float(scores[4]))

    def __eq__(self, other):
        if isinstance(other, EnglishExam):
            return (self.total == other.total) and \
                   (self.listening == other.listening) and \
                   (self.speaking == other.speaking) and \
                   (self.reading == other.reading) and \
                   (self.writing == other.writing)
        else:
            return False

    def __le__(self, other):
        if isinstance(other, EnglishExam):
            return (self.total <= other.total) and \
                   (self.listening <= other.listening) and \
                   (self.speaking <= other.speaking) and \
                   (self.reading <= other.reading) and \
                   (self.writing <= other.writing)
        else:
            return False

    def __ge__(self, other):
        if isinstance(other, EnglishExam):
            return (self.total >= other.total) and \
                   (self.listening >= other.listening) and \
                   (self.speaking >= other.speaking) and \
                   (self.reading >= other.reading) and \
                   (self.writing >= other.writing)
        else:
            return False


    def __repr__(self):
        return self.__class__.__name__ + "({0.total}, {0.listening}, {0.speaking}, {0.reading}, {0.writing})".format(self)


class TOEFL(EnglishExam):
    def __init__(self, total=0, listening=0, speaking=0, reading=0, writing=0):
        super().__init__(total, listening, speaking, reading, writing)


class IELTS(EnglishExam):
    def __init__(self, total=0, listening=0, speaking=0, reading=0, writing=0):
        super().__init__(total, listening, speaking, reading, writing)


class TOEIC(EnglishExam):
    def __init__(self, total=0, listening=0, reading=0):
        super().__init__(total, listening, 0, reading, 0)

    # Input:
    #   'scores':refer to a string of serial scores in format: T/L/R
    #
    def setScores(self, scores):
        scores = scores.split('/')
        if len(scores) == 3:
            self.__init__(float(scores[0]), float(scores[1]), float(scores[2]))

    def __le__(self, other):
        if isinstance(other, TOEIC):
            return (self.total <= other.total) and \
                   (self.listening <= other.listening) and \
                   (self.reading <= other.reading)
        else:
            return False

    def __ge__(self, other):
        if isinstance(other, TOEIC):
            return (self.total >= other.total) and \
                   (self.listening >= other.listening) and \
                   (self.reading >= other.reading)
        else:
            return False

    def __repr__(self):
        return self.__class__.__name__ + "({0.total}, {0.listening}, {0.reading})".format(self)
