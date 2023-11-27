import nltk

class CommentGenerator:
    def __init__(self, student_name, short_name, gender, comment_mapping, student_result, letter_grade):
        self._student_name = student_name
        self._short_name = short_name
        self._gender = gender
        self._comment_mapping = comment_mapping
        self._student_result = student_result
        self._letter_grade = letter_grade

        self._comment_mapping = self._comment_mapping.fillna("")

    def generate_comment(self, probe = False):
        pronoun = "he" if self._gender == "M" else "she"
        adjective = "his" if self._gender == "M" else "her"

        sentence_intro = self._comment_mapping.loc["Intro", self._letter_grade]
        sentence_closing = self._comment_mapping.loc["Closing", self._letter_grade]

        positive_sentences, negative_sentences = self.__collect_comments()
        positive_comments = ""
        negative_comments = ""
        
        if (len(positive_sentences) > 3):
            positive_comments = f"{positive_sentences[0]}. "
            positive_comments += ", ".join(positive_sentences[1:-2])
            positive_comments = " ".join([positive_comments, ", and ", positive_sentences[-2]])
            positive_comments = positive_comments.replace("also", "")
            positive_comments += f". {positive_sentences[-1]}"
        elif (len(positive_sentences) == 3):
            positive_comments = f"{positive_sentences[0]}, and {positive_sentences[1]}"
            positive_comments = positive_comments.replace("also", "")
            positive_comments = ". ".join([positive_comments, positive_sentences[2]])
        elif (len(positive_sentences) == 2):
            positive_comments = f"{positive_sentences[0]}, and {positive_sentences[1]}"
            positive_comments = positive_comments.replace("also", "", 1)
        elif (len(positive_sentences) == 1):
            positive_comments = positive_sentences[0]
            positive_comments = positive_comments.replace("also", "")

        if (len(negative_sentences) > 3):
            negative_comments = f"{negative_sentences[0]}. "
            negative_comments += ", ".join(negative_sentences[1:-2])
            negative_comments = " ".join([negative_comments, ", and ", negative_sentences[-2]])
            negative_comments = negative_comments.replace("also", "")
            negative_comments += f". Further, {negative_sentences[-1]}"
        elif (len(negative_sentences) == 3):
            negative_comments = f"{negative_sentences[0]}, and {negative_sentences[1]}"
            negative_comments = negative_comments.replace("also", "", 1)
            negative_comments = ". ".join([negative_comments, negative_sentences[2]])
        elif (len(negative_sentences) == 2):
            negative_comments = f"{negative_sentences[0]}, and {negative_sentences[1]}"
            negative_comments = negative_comments.replace("also", "", 1)
        elif (len(negative_sentences) == 1):
            negative_comments = negative_sentences[0]
            negative_comments = negative_comments.replace("also", "")

        text = ". ".join([sentence_intro, positive_comments, negative_comments, sentence_closing])

        comment = text.format(short_name = self._short_name, pronoun = pronoun, adjective = adjective)
        comment = self.__format_comment(comment)

        if probe:
            return comment, len(positive_sentences), len(negative_sentences)
        else:
            return comment
        
    def __format_comment(self, comment):
        words = nltk.word_tokenize(comment)

        for word in words:
            if word == "Moreover" or word == "However":
                if words[words.index(word) - 1] == ",":
                    words[words.index(word) - 1] = "."
                elif words[words.index(word) - 1] == "and":
                    words[words.index(word) - 2] = "."
                    words[words.index(word) - 1] = "."

        sentences = " ".join(words)
        sentences = sentences.replace(" .", ".").replace("..", ".").replace(",.", ".").replace(" ,", ",").replace("  ", " ").replace(" !", "!")
        sentences = nltk.sent_tokenize(sentences)
        formatted = [sentence[0].upper() + sentence[1:] for sentence in sentences]
        
        return " ".join(formatted)
    
    def __collect_comments(self):
        positive_sentences = []
        negative_sentences = []
        negative_count = 0

        for key, value in self._student_result.items():
            if key in self._comment_mapping.index:
                comment = self._comment_mapping.loc[key, value]
                if comment != "":
                    if comment.startswith("However"):
                        if negative_count == 0:
                            negative_sentences.append(comment)
                            negative_count += 1
                        else:
                            negative_sentences.append(comment.replace("However,", ""))
                    else:
                        positive_sentences.append(comment)

        return positive_sentences, negative_sentences