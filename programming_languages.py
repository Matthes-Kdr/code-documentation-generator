






class SyntaxVba:
    """
    Muster, die fuer diese Programmiersprache erforderlich sind.
    """


    @staticmethod
    def get_pattern_single_line_comment():
        
        pattern = r"""'    # KommentarApostroph
                .?       # Beliebiges Zeichen
                """
        
        return pattern
        







    @staticmethod
    def get_pattern_end_of_procedure():
        

        # Das folgenden Regex-Muster berücksichtigt nicht das Auskommentieren dieser Zeile
        pattern = r""".*?     # Start mit beliebigen Zeichen
                (?:End)    # Beinhaltet das KEyword
                \s+        # mind. 1 bis n Leerzeichen
                (?:PLACEHOLDER_PROCEDURE_TYPE)    # Beinhaltet das KEyword
                \s+         # mind. 1 bis n Leerzeichen
                .*"""
    


        return pattern
    





    @staticmethod
    def get_pattern_start_of_procedure():
        

        # Das folgenden Regex-Muster berücksichtigt nicht das Auskommentieren dieser Zeile
        pattern = r""".*     # Start mit beliebigen Zeichen
                        (?:Private|Public|Friend)?
                        (?:PLACEHOLDER_PROCEDURE_TYPE)    # Beinhaltet das KEyword
                        \s+        # mind. 1 bis n Leerzeichen
                        (\w+)        # mind. 1 bis n Wortzeichen
                        \(         # Geöffnete Klammer
                        """
    


        return pattern
    



    @staticmethod
    def get_pattern_multiline_comment():
        
        pattern = None

        return pattern







    @staticmethod
    def get_pattern_references(prozedur_name:str):

        patterns = [

            r"(\s{0,}§§___PROC_NAME___§§\b)(?!\s{0,}=)".replace("§§___PROC_NAME___§§", prozedur_name),

        ]

        return patterns




