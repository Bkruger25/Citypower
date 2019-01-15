Module old
    ' Create XmlWriterSettings.
    Dim settings As XmlWriterSettings = New XmlWriterSettings()
        settings.Indent = True
    ' Create XmlWriter.   

      Using updateWriter As XmlWriter = XmlWriter.Create("C:\Users\bkruger\Desktop\updateCase.xml", settings)
    'begin writing
            updateWriter.WriteStartDocument()
    'root
            updateWriter.WriteStartElement("BizAgiWSParam")
            updateWriter.WriteElementString("Domain", "Domain")
            updateWriter.WriteElementString("Username", "Admon")
            updateWriter.WriteStartElement("Cases")
            updateWriter.WriteStartElement("Case")
            updateWriter.WriteElementString("Process", "ServiceConnectionApp")
            updateWriter.WriteStartElement("Entities")
            updateWriter.WriteStartElement("ServiceConnectionApp")
            updateWriter.WriteElementString("NotificationNo", notificationNo(i - 1))
            updateWriter.WriteElementString("NotificationDate", notificationDate(i - 1))
            updateWriter.WriteElementString("NotificationDesc", notificationDesc(i - 1))
            updateWriter.WriteElementString("SGID", sgID(i - 1))
            updateWriter.WriteElementString("Township", township(i - 1))
            updateWriter.WriteElementString("FunctionalLocation", functionalLoc(i - 1))
            updateWriter.WriteElementString("Account", account(i - 1))
            updateWriter.WriteElementString("Name", name(i - 1))
            updateWriter.WriteElementString("CustomerTown", customerTown(i - 1))
            updateWriter.WriteElementString("StreetNo", streetNo(i - 1))
            updateWriter.WriteElementString("PostCode", postcode(i - 1))
            updateWriter.WriteElementString("StreetAddress", streetAddress(i - 1))
            updateWriter.WriteEndElement()
            updateWriter.WriteEndElement()
            updateWriter.WriteEndDocument()
        End Using

    'create new case writer xml file
        Using newWriter As XmlWriter = XmlWriter.Create("C:\Users\bkruger\Desktop\newCase.xml", settings)
    'begin writing
            newWriter.WriteStartDocument()
    'root
            newWriter.WriteStartElement("BizAgiWSParam")
            newWriter.WriteElementString("domain", "domain")
            newWriter.WriteElementString("userName", "admon")
            newWriter.WriteStartElement("Cases")
    'write values to xml
    ''newWriter.WriteStartElement("Case")
            newWriter.WriteElementString("Process", "ServiceConnectionApp")
            newWriter.WriteStartElement("Entities")
            newWriter.WriteStartElement("ServiceConnectionApp")
    ''newWriter.WriteAttributeString("businessKey", "NotificationNo ='" & notificationNo(i - 1) & "'")
            newWriter.WriteElementString("NotificationNo", notificationNo(i - 1))
            newWriter.WriteElementString("NotificationDate", notificationDate(i - 1))
            newWriter.WriteElementString("NotificationDesc", notificationDesc(i - 1))
            newWriter.WriteElementString("SGID", sgID(i - 1))
            newWriter.WriteElementString("Township", township(i - 1))
            newWriter.WriteElementString("FunctionalLocation", functionalLoc(i - 1))
            newWriter.WriteElementString("Account", account(i - 1))
            newWriter.WriteElementString("Name", name(i - 1))
            newWriter.WriteElementString("CustomerTown", customerTown(i - 1))
            newWriter.WriteElementString("StreetNo", streetNo(i - 1))
            newWriter.WriteElementString("PostCode", postcode(i - 1))
            newWriter.WriteElementString("StreetAddress", streetAddress(i - 1))
            newWriter.WriteEndElement()
            newWriter.WriteEndElement()
            newWriter.WriteEndElement()
            newWriter.WriteEndElement()
            newWriter.WriteEndDocument()
        End Using
End Module
