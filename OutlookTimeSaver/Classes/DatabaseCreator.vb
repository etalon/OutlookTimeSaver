Public Class DatabaseCreator

    ''' <summary>
    ''' Initialisierung der Datenbank und eventuelle Durchführung von Updates
    ''' </summary>
    ''' <remarks></remarks>
    Public Shared Sub Init()

        Dim hasCurrentDbVersion As Boolean = False
        Dim currentDbVersion As Integer

        DatabaseWrapper.Parameter.DefaultDatabaseWrapperConfigOnNewConfigFile = New DatabaseWrapper._EtiDatabaseWrapperConfig
        With DatabaseWrapper.Parameter.DefaultDatabaseWrapperConfigOnNewConfigFile
            .DbType = "Sqlite"
            .EnableSqlLogging = True
            With .SqliteConfig
                .FileName = "OutlookTimeSaverDb.sqlite"
            End With
        End With


        Using db As DatabaseWrapper = DatabaseWrapper.CreateInstance()
            With db

                If Not .GetTables.Contains("config") Then
                    .ExecuteNonQuery("CREATE TABLE config (id nvarchar(50), content nvarchar(50), PRIMARY KEY(id))")
                    .ExecuteNonQuery("INSERT INTO config (id, content) VALUES ('DB_VERSION','0')")
                End If

                currentDbVersion = .ExecuteScalar(Of Integer)("SELECT content FROM config WHERE id = @0", "DB_VERSION")

                Select Case currentDbVersion
                    Case 0

                        .ExecuteNonQuery("CREATE TABLE firstname (name nvarchar(50), gender nvarchar(1), PRIMARY KEY(name))")
                        insertFirstNames(db)

                    Case 1
                        .ExecuteNonQuery("CREATE TABLE salutation (recipients nvarchar(200) NOT NULL, text nvarchar(100) NOT NULL, PRIMARY KEY(recipients));")

                    Case 2
                        .ExecuteNonQuery("CREATE TABLE exchangeaddress (address nvarchar(300) NOT NULL, email nvarchar(100) NOT NULL, PRIMARY KEY(address));")

                    Case 3
                        .ExecuteNonQuery("ALTER TABLE salutation ADD COLUMN mailcount integer NOT NULL DEFAULT 0; ALTER TABLE salutation ADD COLUMN  firstname nvarchar(50) NOT NULL DEFAULT ''; ALTER TABLE salutation ADD COLUMN lastname nvarchar(50) NOT NULL DEFAULT ''; ALTER TABLE salutation ADD COLUMN shownname nvarchar(100) NOT NULL DEFAULT ''")

                    Case 4
                        .ExecuteNonQuery("ALTER TABLE salutation ADD COLUMN gender nvarchar(1) NOT NULL DEFAULT ''")

                    Case 5
                        .ExecuteNonQuery("CREATE TABLE recipient (email nvarchar (200) NOT NULL, salutation nvarchar (100) NOT NULL, mailcount integer NOT NULL DEFAULT 0, firstname nvarchar (50) NOT NULL DEFAULT '', lastname nvarchar (50) NOT NULL DEFAULT '', displayname nvarchar (100) NOT NULL DEFAULT '', gender nvarchar(1) NOT NULL DEFAULT '', PRIMARY KEY (email));")
                        .ExecuteNonQuery("INSERT INTO recipient(email,salutation,mailcount,firstname,lastname,shownname,gender) SELECT recipients,text,0,'','','','' FROM salutation;")
                        .ExecuteNonQuery("DROP TABLE salutation")

                    Case Else
                        ' Kein Update notwendig
                        hasCurrentDbVersion = True
                End Select

                If Not hasCurrentDbVersion Then
                    ' Datenbankversion um eins erhöhen
                    updateDbVersion(db, currentDbVersion + 1)
                End If

            End With
        End Using

        If Not hasCurrentDbVersion Then
            ' Achtung: Rekursiver Aufruf, damit die nächste Datenbankversion installiert wird
            Init()
            Return
        End If

    End Sub

    Private Shared Sub insertFirstNames(p_Db As DatabaseWrapper)

        Dim r() As String
        Dim sql = "INSERT INTO firstname (name, gender) VALUES (@0,@1)"

        p_Db.StartTransaction()

        For Each l In Split(ReadTextFileFromResources("OutlookTimeSaver.firstnames.txt"), vbCrLf)
            r = l.Split(vbTab.Chars(0))
            p_Db.ExecuteNonQuery(sql, r(0), r(1))
        Next

        p_Db.CommitTransaction()

    End Sub

    Private Shared Sub updateDbVersion(ByRef db As DatabaseWrapper, ByVal newVersion As Integer)

        db.ExecuteNonQuery("UPDATE config Set content = @0 WHERE id = @1", newVersion, "DB_VERSION")

    End Sub

End Class
