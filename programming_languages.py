






class SyntaxVba:
    """
    Muster, die fuer diese Programmiersprache erforderlich sind.
    """


    @staticmethod
    def get_pattern_single_line_comment():
        
        # TODO: ...
        # pattern_single_line_comment = re.compile(__regex_ausschlus_kommentar_pattern, re.VERBOSE | re.IGNORECASE)
        pass


    @staticmethod
    def get_pattern_multiline_comment():
        
        pattern_multiline_comment = None

        return pattern_multiline_comment





    @staticmethod
    def get_pattern_references(prozedur_name:str):

        pattern_references = [

            r"(\s{0,}§§___PROC_NAME___§§\b)(?!\s{0,}=)".replace("§§___PROC_NAME___§§", prozedur_name),

        ]

        return pattern_references




