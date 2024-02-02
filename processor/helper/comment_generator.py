import google.generativeai as genai
import nltk
from termcolor import colored

import config

class CommentGenerator:
    """
    Generates a comment based on the student's result and the comment mapping.

    Attributes:
        student_name (str): The student's name.
        short_name (str): The student's short name.
        gender (str): The student's gender.
        comment_mapping (pandas.DataFrame): The comment mapping.
        student_result (dict): The student's result.
        letter_grade (str): The student's letter grade.
    
    Methods:
        Public:
        generate_comment(self, probe = False, autocorrect = False, ltm = None): Generates a comment based on the student's result and the comment mapping.
        
        Private:
        __format_comment(self, comment): Formats the comment.
        __collect_comments(self): Collects the comments from the comment mapping.
        __assemble_positive_comments(self, positive_sentences): Assembles the positive comments.
        __assemble_negative_comments(self, negative_sentences): Assembles the negative comments.
    """

    def __init__(self, student_name, short_name, gender, comment_mapping, student_result, letter_grade):
        """
        Initialize the CommentGenerator instance.
        
        Args:
            Follows the class attributes.
        """
        self._student_name = student_name
        self._short_name = short_name
        self._gender = gender
        self._comment_mapping = comment_mapping
        self._student_result = student_result
        self._letter_grade = letter_grade

        self._comment_mapping = self._comment_mapping.fillna("")

    # TODO: Change return type to proper dictionary type ?
    def generate_comment(self, probe = False, autocorrect = False, ltm = None):
        """
        Generates a comment based on the student's result and the comment mapping.

        Args:
            probe (bool): Whether the caller want's to 'probe' into the process by returning the number of positive and negative comments.
            autocorrect (bool): Whether to autocorrect the comment using LanguageTool.
            ltm (module): The LanguageTool Instance Manager module. This is to inject the module to the class. 
                          If not specified, the class will try to import the module itself
        
        Returns:
            str: The generated comment.
            int: The number of positive comments. (Only if probe is True)
            int: The number of negative comments. (Only if probe is True)

        Raises:
            ImportError: If the LanguageTool module is not installed.
        """
        if ltm is None:
            try:
                import processor.helper.language_tool_master as ltm
            except ImportError:
                raise ImportError("LanguageTool is not installed. Please install LanguageTool to use this feature.")

        pronoun = "he" if self._gender == "M" else "she"
        adjective = "his" if self._gender == "M" else "her"

        sentence_intro = self._comment_mapping.loc["Intro", self._letter_grade] if self._letter_grade != " " else ""
        sentence_closing = self._comment_mapping.loc["Closing", self._letter_grade] if self._letter_grade != " " else ""

        positive_sentences, negative_sentences = self.__collect_comments()
        positive_text = self.__assemble_positive_comments(positive_sentences)
        negative_text = self.__assemble_negative_comments(negative_sentences)

        text = ". ".join([sentence_intro, positive_text, negative_text, sentence_closing])

        comment = text.format(short_name = "VDC", pronoun = pronoun, adjective = adjective)
        if autocorrect:
            tool = ltm.get_tool()
            if tool.check(comment):
                comment = tool.correct(comment)
                
        comment = comment.replace("VDC", self._short_name)
        comment = self.__format_comment(comment)

        if probe:
            return comment, len(positive_sentences), len(negative_sentences)
        else:
            return comment
        
    def __format_comment(self, comment):
        """
        Formats the comment by capitalizing the first letter of each sentence and removing unnecessary spaces;
        Combines sentences that are separated by conjunctions;
        Fixes conjunctions that are placed after commas.
        
        Args:
            comment (str): The comment to format.

        Returns:
            str: The formatted comment.
        """
        words = nltk.word_tokenize(text = comment, language = "english")

        conjunctions = ["Moreover", "However", "Further", "Also", "Besides", "Additionally", "Furthermore", "In addition", "In addition to", "In addition", "Though", "On the other hand"]

        for word in words:
            if word in conjunctions:
                if words[words.index(word) - 1] == ",":
                    words[words.index(word) - 1] = "."
                elif words[words.index(word) - 1] == "and":
                    words[words.index(word) - 2] = "."
                    words[words.index(word) - 1] = "."

        sentences = " ".join(words)
        sentences = sentences.replace(" .", ".").replace("..", ".").replace(",.", ".").replace(" ,", ",").replace("  ", " ").replace(" !", "!").replace(" '", "'").replace(".,", ".")
        sentences = nltk.sent_tokenize(text = sentences, language = "english")
        formatted = [sentence[0].upper() + sentence[1:] for sentence in sentences]
        
        return " ".join(formatted)
    
    def __collect_comments(self):
        """
        Collects the comments from the comment mapping.
        
        The function will collect the comments from the comment mapping based on the student's result.
        It will also separate the comments into positive and negative comments by recognizing the existence
        of the word "However" in the comments.

        Returns:
            list: The list of positive comments.
            list: The list of negative comments.
        """
        positive_sentences = []
        negative_sentences = []
        negative_count = 0

        for key, value in self._student_result.items():
            if key in self._comment_mapping.index and value in self._comment_mapping.columns:
                comment = self._comment_mapping.loc[key, value]
                if len(comment) > 0 and comment != "nan":
                    if comment.startswith("However"):
                        if negative_count == 0:
                            negative_sentences.append(comment)
                            negative_count += 1
                        else:
                            negative_sentences.append(comment.replace("However,", ""))
                    else:
                        positive_sentences.append(comment)
            else:
                print(colored(f"Info: Grade '{value}' or Goal '{key}' is not found in the comment mapping possibly due to incomplete skills and assessment grading. Skipping this comment.", "yellow"))

        return positive_sentences, negative_sentences
    
    def __assemble_positive_comments(self, positive_sentences):
        """
        Assembles the positive comments.

        The function will assemble the positive comments into a single string.
        It will also remove unnecessary occurrences of the word "also" in the comments.

        Args:
            positive_sentences (list): The list of positive comments.

        Returns:
            str: The assembled positive comments.
        """
        positive_comments = ""

        if (len(positive_sentences) > 5):
            positive_comments = f"{positive_sentences[0]}, and {positive_sentences[1]}. "
            remaining_sentences = len(positive_sentences) - 3
            for i in range(2, len(positive_sentences) - 2):
                if remaining_sentences == 3:
                    positive_comments = positive_comments.replace("also", "")
                    positive_comments += f". In addition, {positive_sentences[i]}. "
                    positive_comments += f"{positive_sentences[i+1]}, and {positive_sentences[i + 2].replace('also', '')}."
                    break
                elif remaining_sentences == 2:
                    positive_comments = positive_comments.replace("also", "")
                    positive_comments += f". In addition, {positive_sentences[i]}, and {positive_sentences[i + 1]}."
                    break
                elif remaining_sentences == 1:
                    positive_comments = positive_comments.replace("also", "")
                    positive_comments += f". {positive_sentences[i]}."
                    break
                elif i % 3 == 1:
                    positive_comments = " ".join([positive_comments, ", and ", positive_sentences[i]])
                    remaining_sentences -= 1
                elif remaining_sentences > 3:
                    positive_comments += f", {positive_sentences[i]}"
                    remaining_sentences -= 1
            positive_comments += f". {positive_sentences[-1]}"
        elif (len(positive_sentences) > 3):
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
        
        return positive_comments
    
    def __assemble_negative_comments(self, negative_sentences):
        """
        Assembles the negative comments.

        The function will assemble the negative comments into a single string.
        It will also remove unnecessary occurrences of the word "also" in the comments.

        Args:
            negative_sentences (list): The list of negative comments.

        Returns:
            str: The assembled negative comments.
        """
        negative_comments = ""

        if (len(negative_sentences) > 5):
            negative_comments = f"{negative_sentences[0]}, and {negative_sentences[1]}. "
            remaining_sentences = len(negative_sentences) - 3
            for i in range(2, len(negative_sentences) - 2):
                if remaining_sentences == 3:
                    negative_comments = negative_comments.replace("also", "")
                    negative_comments += f". Besides, {negative_sentences[i]}. "
                    negative_comments += f"{negative_sentences[i+1]}, and {negative_sentences[i + 2].replace('also', '')}."
                    break
                elif remaining_sentences == 2:
                    negative_comments = negative_comments.replace("also", "")
                    negative_comments += f". Besides, {negative_sentences[i]}, and {negative_sentences[i + 1]}."
                    break
                elif remaining_sentences == 1:
                    negative_comments = negative_comments.replace("also", "")
                    negative_comments += f". {negative_sentences[i]}."
                    break
                elif i % 3 == 1:
                    negative_comments = " ".join([negative_comments, ", and ", negative_sentences[i]])
                    remaining_sentences -= 1
                elif remaining_sentences > 3:
                    negative_comments += f", {negative_sentences[i]}"
                    remaining_sentences -= 1
            negative_comments += f". {negative_sentences[-1]}"
        elif (len(negative_sentences) > 3):
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

        return negative_comments
    
class AICommentGenerator:
    """
    Generates a comment based on the student's result using AI.

    Attributes:
        None

    Methods:
        generate_comment(self, nickname, gender, result, verbose = False): Generates a comment based on the student's result using AI.
        rephrase(self, source): Rephrases a pre-generated comment using AI.
    """
    genai.configure(api_key = config.get_config("genai_api_key"))

    config = genai.GenerationConfig(candidate_count = 1, temperature = 0.2, max_output_tokens = 200, top_k = 20, top_p = 0.8)
    model = genai.GenerativeModel('gemini-pro')

    generation_config = {
    "temperature": 0.2,
    "max_output_tokens": 200,
    "top_k": 20,
    "top_p": 0.8,
    }

    safety = [
        {
            "category": "HARM_CATEGORY_HATE_SPEECH",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        },
        {
            "category": "HARM_CATEGORY_HARASSMENT",
            "threshold": "BLOCK_MEDIUM_AND_ABOVE"
        }
    ]

    def generate_comment(self, nickname, gender, final_grade, result, verbose = False):
        """
        Generates a comment based on the student's result using AI.
        
        Args:
            nickname (str): The student's nickname.
            gender (str): The student's gender.
            result (dict): The student's result.
            verbose (bool): Whether to print the prompt and response.
        
        Returns:
            str: The generated comment.
        """
        assembled_result = ""

        for goal, grade in result.items():
            assembled_result += f"{goal}: {grade}\n"

        if gender == "M":
            gender_normalized = "male"
        elif gender == "F":
            gender_normalized = "female"

        base_prompt = "Write a simple, single paragraph report card comment for a student according to their grades for each goal. Grade A is best, B is better, C is okay, and Grade D being worst. Make it between 30 and 70 words. Strictly no more than 70 words. Don't forget to include a simple opening and a closing. Be truthful and open about the comment even if it is blatant. Do not mention the grades in the commentary."
        parametric_prompt = f"The student's nickname is {nickname}. This student is a {gender_normalized} and achieved an overall grade of {final_grade}.\nGoals and grades for each goal:\n{assembled_result}"
        response = self.model.generate_content(f"{base_prompt} {parametric_prompt}", safety_settings = self.safety, generation_config = self.generation_config)

        if verbose:
            print(f"\nPrompt: {base_prompt} {parametric_prompt}\n")
            print(f"\nCandidates: {response.candidates}")
            print(f"\nResponse: {response.text}")
            print(f"\nResponse Length: {len(response.text.split())} words\n")

        if len(response.text.split()) > 80:
            if verbose:
                print(colored("(!) Response too long. Rephrasing the response.", "light_cyan"))
            
            return self.rephrase(response.text)

        return response.text

    def rephrase(self, source):
        """
        Rephrases a pre-generated comment using AI.
        
        Args:
            source (str): The source comment to rephrase.
            
        Returns:
            str: The rephrased comment.
        """
        response = self.model.generate_content(f"Shorten the following content. Use simple english and do not add personal opinions. The content should be less than 75 words. Content: {source}.", 
                                               safety_settings = self.safety, 
                                               generation_config = self.generation_config)
        return response.text