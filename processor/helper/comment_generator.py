import nltk

class CommentGenerator:
    def __init__(self, student_name, short_name, gender):
        self._student_name = student_name
        self._short_name = short_name
        self._gender = gender

    def generate_comment(self):
        pronoun = "he" if self._gender == "M" else "she"
        adjective = "his" if self._gender == "M" else "her"

        text = "{short_name} is an attentive and respectful student. {pronoun} has good art expression and is good at applying colour to {adjective} artwork. {pronoun} is also able to use art to express {adjective} ideas and feelings. Moreover, {pronoun} has a creative way of using art media. Keep up your good work, {short_name}!"

        comment = text.format(short_name = self._short_name, pronoun = pronoun, adjective = adjective)
        comment = self.__format_comment(comment)

        return comment
        
    def __format_comment(self, comment):
        sentences = nltk.sent_tokenize(comment)
        formatted = [sentence[0].upper() + sentence[1:] for sentence in sentences]
        return " ".join(formatted)