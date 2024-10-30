import re
import time

import google.generativeai as genai
import nltk
from datetime import datetime
from termcolor import colored
from google.ai.generativelanguage import Candidate
from google.api_core.exceptions import ResourceExhausted

import components.report_generator.manifest as manifest
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
        generate_comment(self, probe = False, autocorrect = False): Generates a comment based on the student's result and the comment mapping.
        
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
    def generate_comment(self, probe = False, autocorrect = False):
        """
        Generates a comment based on the student's result and the comment mapping.

        Args:
            probe (bool): Whether the caller want's to 'probe' into the process by returning the number of positive and negative comments.
            autocorrect (bool): Whether to autocorrect the comment using LanguageTool.
        
        Returns:
            str: The generated comment.
            int: The number of positive comments. (Only if probe is True)
            int: The number of negative comments. (Only if probe is True)

        Raises:
            ImportError: If the LanguageTool module is not installed.
        """
        pronoun = "he" if self._gender == "M" else "she"
        adjective = "his" if self._gender == "M" else "her"

        sentence_intro = self._comment_mapping.loc["Intro", self._letter_grade] if self._letter_grade != " " else ""
        sentence_closing = self._comment_mapping.loc["Closing", self._letter_grade] if self._letter_grade != " " else ""

        positive_sentences, negative_sentences = self.__collect_comments()
        positive_text = self.__assemble_positive_comments(positive_sentences)
        negative_text = self.__assemble_negative_comments(negative_sentences)

        text = ". ".join([sentence_intro, positive_text, negative_text, sentence_closing])

        comment = text.format(short_name = "VDC", pronoun = pronoun, adjective = adjective)
        print(f"\nGenerated Comment:\n{comment}")
        if autocorrect:
            ai = AICommentGenerator(None)
            comment = ai.grammar_check(comment)
            print(colored("\n[Comment autocorrected using AI]", "yellow"))
            print(f"Autocorrected Comment:\n{comment}\n")
                
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
        config (genai.GenerationConfig): The Gen AI generation configuration.
        model (genai.GenerativeModel): The Gen AI generative model.
        manifest (manifest.Manifest): The manifest instance for logging.

    Methods:
        get_base_prompt(self): Gets the base prompt from a predefined file.
        generate_comment(self, nickname, gender, result, verbose = False): Generates a comment based on the student's result using AI.
        rephrase(self, source): Rephrases a pre-generated comment using AI.
        grammar_check(self, text): Checks the grammar of a comment using Gemini AI.
    """

    def __init__(self, manifest: manifest.Manifest):
        """
        Initialize the AICommentGenerator instance.
        """
        genai.configure(api_key = config.get_config("genai_api_key"))

        self.config = genai.GenerationConfig(candidate_count = 1, temperature = 0.2, max_output_tokens = 225, top_k = 20, top_p = 0.8)
        self.model = genai.GenerativeModel('gemini-1.5-flash')
        self.manifest = manifest

        self.safety = [
            {
                "category": "HARM_CATEGORY_HATE_SPEECH",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            },
            {
                "category": "HARM_CATEGORY_HARASSMENT",
                "threshold": "BLOCK_MEDIUM_AND_ABOVE"
            }
        ]

    def get_base_prompt(self):
        """
        Gets the base prompt from a predefined file.

        Returns:
            str: The base prompt.
        """
        with open("ai/base_prompt.txt", "r") as file:
            base_prompt = file.read()
            return base_prompt

    def generate_comment(self, nickname, gender, final_grade, sna_list, verbose = False, retry_count = 0):
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
        rephrase = False
        assembled_result = ""
        finish_time = ""

        for goal, grade in sna_list.items():
            assembled_result += f"{goal}: {grade}\n"

        if gender == "M":
            gender_normalized = "male"
        elif gender == "F":
            gender_normalized = "female"

        # Dynamic length calculation
        goals_counts = len(sna_list)
        max_length = 0

        match goals_counts:
            case 3:
                max_length = 800
            case 4:
                max_length = 700
            case 5:
                max_length = 600
            case 6:
                max_length = 500
            case 7 | 8 | 9 | 10:
                max_length = 400
            case _:
                max_length = 900

        base_prompt = self.get_base_prompt()
        parametric_prompt = f"The student's nickname is {nickname}. This student is a {gender_normalized} and achieved an overall grade of {final_grade}.\nGoals and grades for each goal:\n{assembled_result}"
        final_prompt = f"{base_prompt}\nMake it between 400 and {max_length} characters. Strictly no more than {max_length} characters.\n{parametric_prompt}"

        if verbose:
            print(f"\nPrompt: {final_prompt}\n")

        try:
            response = self.model.generate_content(f"{final_prompt}", safety_settings = self.safety, generation_config = self.config)
        except ResourceExhausted as e:
            print(colored(f"(!) Error: Resource Exhausted. Retrying...", "red"))
            if retry_count < 3:
                print(colored(f"(i) Retry Count: {retry_count + 1} out of 3. Waiting for 30 seconds before retrying...", "yellow"))
                time.sleep(30)
                return self.generate_comment(nickname, gender, final_grade, sna_list, verbose = verbose, retry_count = retry_count + 1)
        except Exception as e:
            print(type(e))
            print(colored(f"(!) Error: {e}", "red"))
            return f"AI Comment Generation Error! Reason: {e}\nPlease regenerate report for this student manually."

        if response.candidates[0].finish_reason is not Candidate.FinishReason.STOP:
            print(colored(f"(!) Server stopped because: {response.candidates[0].finish_reason.name}. Aborting.\n", "light_cyan"))
            return f"AI Comment Generation Error! Reason: Server stopped because: {response.candidates[0].finish_reason.name}."

        if verbose:
            print(f"\nCandidates: {response.candidates}")
            print(f"\nResponse: {response.text}")
            print(f"\nResponse Length (Chars): {len(response.text)} characters")
            print(f"Response Length (Words): {len(response.text.split())} words\n")
            print(f"Input Tokens Used: {response.usage_metadata.prompt_token_count}")
            print(f"Output Tokens Used: {response.usage_metadata.candidates_token_count}")

        if len(response.text) > max_length:
            if verbose:
                print(colored("(!) Response too long. Rephrasing the response.", "light_cyan"))
            
            rephrase = True
            final_response, rephrase_metadata = self.rephrase(response.text, max_length = max_length)

            if verbose:
                print(f"\nRephrased Response: {final_response}\n")
                print(f"Rephrased Response Length (Chars): {len(final_response)} characters")
                print(f"Rephrased Response Length (Words): {len(final_response.split())} words\n")
                print(f"Rephrased Response Tokens Used: {rephrase_metadata.prompt_token_count}")
                print(f"Rephrased Response Tokens Used: {rephrase_metadata.candidates_token_count}")

            finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")

        # Remove unnecessary new lines if not rephrasing
        if not rephrase:
            split_response = response.text.split("\n")
            final_response = " ".join(split_response)
            final_response = final_response.replace("  ", " ").replace(" .", ".").replace(" ,", ",").replace(" !", "!").replace(" '", "'").replace(".,", ".")
            finish_time = datetime.now().strftime("%Y-%m-%d %H:%M:%S")
            input_tokens = response.usage_metadata.prompt_token_count
            output_tokens = response.usage_metadata.candidates_token_count
        else:
            input_tokens = rephrase_metadata.prompt_token_count
            output_tokens = rephrase_metadata.candidates_token_count
        
        self.manifest.add_entry(student = nickname, 
                                comment_orig = final_response if not rephrase else response.text, 
                                comment_short = final_response if rephrase else "-", 
                                length_chars = len(final_response),
                                length_words = len(final_response.split()),
                                input_tokens = input_tokens,
                                output_tokens = output_tokens,
                                status = "Done", 
                                completed_at = finish_time,
                                error = None if not rephrase else "Max Length Exceeded. Rephrased.")
        
        if verbose:
            print("Fixing gender strings if necessary...")

        final_response = self.__fix_gender_strings(gender, final_response)

        return final_response

    def rephrase(self, source, max_length = 600, retry_count = 0):
        """
        Rephrases a pre-generated comment using AI.
        
        Args:
            source (str): The source comment to rephrase.
            
        Returns:
            str: The rephrased comment.
        """
        try:
            response = self.model.generate_content(f"Shorten the following content. Use simple english and do not add personal opinions. The content should be less than {max_length} characters BUT DO NOT WRITE LESS THAN 400 CHARACTERS. Content: {source}.", 
                                                   safety_settings = self.safety, 
                                                   generation_config = self.config)
        except ResourceExhausted as e:
            print(colored(f"(!) Error: Resource Exhausted. Retrying...", "red"))
            if retry_count < 3:
                print(colored(f"(i) Retry Count: {retry_count + 1} out of 3. Waiting for 30 seconds before retrying...", "yellow"))
                time.sleep(30)
                return self.rephrase(source, max_length = max_length, retry_count = retry_count + 1)
        except Exception as e:
            print(colored(f"(!) Error: {e}", "red"))
            return f"AI Comment Generation Error! Reason: {e}\nPlease regenerate report for this student manually."

        # Remove unnecessary new lines
        split_response = response.text.split("\n")
        final_response = " ".join(split_response)
        final_response = final_response.replace("  ", " ").replace(" .", ".").replace(" ,", ",").replace(" !", "!").replace(" '", "'").replace(".,", ".")

        return final_response, response.usage_metadata
    
    def grammar_check(self, text):
        """
        Checks the grammar of a comment using Gemini AI.
        
        Args:
            text (str): The comment to check.
            
        Returns:
            str: The corrected comment.
        """
        try:
            response = self.model.generate_content(f"You are a professional writer. You are reviewing a report for a student made by a machine. You want to make sure that the report is grammatically correct. The report is: {text}.",
                                                   safety_settings = self.safety,
                                                   generation_config = self.config)
        except Exception as e:
            print(colored(f"(!) Error: {e}", "red"))
            return f"AI Grammar Checking Error! Reason: {e}\nPlease regenerate report for this student manually."
        
        # Remove unnecessary new lines
        split_response = response.text.split("\n")
        final_response = " ".join(split_response)
        final_response = final_response.replace("  ", " ").replace(" .", ".").replace(" ,", ",").replace(" !", "!").replace(" '", "'").replace(".,", ".")

        return final_response
    
    def __fix_gender_strings(self, gender, comment):
        """
        Checks the pronouns, and adjectives in a comment based on
        the student's gender as provided.

        Args:
            gender (str): The student's gender. Use ```M``` for male and ```F``` for female.
            comment (str): The comment to check.

        Returns:
            str: The updated comment.
        """
        if gender not in ["M", "F"]:
            raise ValueError

        patterns = { 
            # (old, new)
            "M": [("she", "he"), ("her", "his")],
            "F": [("he", "she"), ("his", "her")]
        }

        for old, new in patterns[gender]:
            comment = re.sub(rf"\b{old}\b", new, comment, flags = re.IGNORECASE)
            comment = re.sub(rf"\b{old.capitalize()}\b", new.capitalize(), comment)

        return comment