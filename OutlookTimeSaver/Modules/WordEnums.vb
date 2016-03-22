Module WordEnums

    Public Enum WDUnits

        wdCell = 12 ' A cell.
        wdCharacter = 1 'A character.
        wdCharacterFormatting = 13  'Character formatting.
        wdColumn = 9    'A column.
        wdItem = 16 'The selected item.
        wdLine = 5  'A line.
        wdParagraph = 4 'A paragraph.
        wdParagraphFormatting = 14  'Paragraph formatting.
        wdRow = 10  'A row.
        wdScreen = 7    'The screen dimensions.
        wdSection = 8   'A section.
        wdSentence = 3  'A sentence.
        wdStory = 6 'A story.
        wdTable = 15    'A table.
        wdWindow = 11   'A window.
        wdWord = 2  'A word.

    End Enum

    Public Enum WDMovementType
        wdExtend = 1    ' The End Of the selection Is extended To the End Of the specified unit.
        wdMove = 0  ' The selection Is collapsed To an insertion point And moved To the End Of the specified unit. Default.
    End Enum

End Module
