package tnefdecoder

import (
	"vcard"
//	"strings"
//	"bytes"
	//"fmt"
	"strconv"
	b64 "encoding/base64"
)

func ExtractVCard(t *TnefObject, vc vcard.IVCard) {
	var (
		attr *Attribute
		attrValue string
	)

	/*
		Importing to Contact Objects
		The VERSION type is not imported to the Contact object.
		Exporting from Contact Objects
		The VERSION type is set to 3.0.
	*/
/*
	for _, ta := range t.Attributes {
		fmt.Printf("\r\nTag: %d ( %X ) Value: %s",ta.Name,ta.Name, string(ta.Data))
	}
*/

	/*
	 FN -> PidTagDisplayName || PidTagNormalizedSubject || PidTagConversationTopic
	 The FN type is generated from either the PidTagDisplayName or PidTagNormalizedSubject property. If both of these properties are set, the PidTagDisplayName property is used.
	*/

	attr = t.GetAttribute(MapiPidTagDisplayName, "mapi"); //PidTagDisplayName
	attrValue = ""
	if attr == nil {
		attr = t.GetAttribute(MapiPidTagNormalizedSubject, "mapi"); // PidTagNormalizedSubject
		if (attr == nil) {
			attr = t.GetAttribute(MapiPidTagConversationTopic, "mapi"); // PidTagConversationTopic
		}
	}
	if (attr != nil) {
		attrValue = attr.GetStringValue()
	}
	pFnValue := vcard.NewText(attrValue)

	if !pFnValue.IsEmpty() {
		propFn := vc.CreateProperty("fn")
		propFn.AddValue(pFnValue)

		vc.AddProperty(propFn)
	}

	/*
	 N -> PidTagSurname, PidTagGivenName, PidTagMiddleName, PidTagDisplayNamePrefix (Honorific Prefixes), PidTagGeneration (Honorific Postfixes)
	*/

	// family name
	attr = t.GetAttribute(MapiPidTagSurname, "mapi");
	attrFamilyValue := ""
	if attr != nil {
		attrFamilyValue = attr.GetStringValue()
	}

	// given
	attr = t.GetAttribute(MapiPidTagGivenName, "mapi");
	attrGivenValue := ""
	if attr != nil {
		attrGivenValue = attr.GetStringValue()
	}

	// middle name
	attr = t.GetAttribute(MapiPidTagMiddleName, "mapi");
	attrMiddleValue := ""
	if attr != nil {
		attrMiddleValue = attr.GetStringValue()
	}

	// prefix
	attr = t.GetAttribute(MapiPidTagDisplayNamePrefix, "mapi");
	attrPrefixValue := ""
	if attr != nil {
		attrPrefixValue = attr.GetStringValue()
	}

	// suffix
	attr = t.GetAttribute(MapiPidTagGeneration, "mapi");
	attrSuffixValue := ""
	if attr != nil {
		attrSuffixValue = attr.GetStringValue()
	}

	propNValue := vcard.NewName()
	propNValue.AddFamilyName(attrFamilyValue)
	propNValue.AddGivenName(attrGivenValue)
	propNValue.AddMiddleName(attrMiddleValue)
	propNValue.AddHonorificPrefix(attrPrefixValue)
	propNValue.AddHonorificSuffix(attrSuffixValue)

	if !propNValue.IsEmpty() {
		propN := vc.CreateProperty("n")
		propN.AddValue(propNValue)
		vc.AddProperty(propN)
	}

	// nickname
	attr = t.GetAttribute(MapiPidTagNickname, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	propNicknameValue := vcard.NewText(attrValue)

	if !propNicknameValue.IsEmpty() {
		propNickname := vc.CreateProperty("nickname")
		propNickname.AddValue(propNicknameValue)
		vc.AddProperty(propNickname)
	}


	for _, att := range t.Attachments {
		hasPhotoAttr := att.GetAttribute(MapiPidTagAttachmentContactPhoto, "mapi");

		if hasPhotoAttr!= nil && hasPhotoAttr.GetBoolValue() {
			// the contact has picture
			photoValue := vcard.NewPhoto(b64.StdEncoding.EncodeToString(att.Data))
			photoValue.IsB64Encoded = true
			photoProp := vc.CreateProperty("photo")
			photoProp.AddValue(photoValue)

			vc.AddPropertyParameter(photoProp, "ENCODING", []string{"base64"})
			vc.AddPropertyParameter(photoProp, "TYPE", []string{"JPEG"})

			vc.AddProperty(photoProp)
			break
		}
	}



	// birthday
	// vCard data format: BDAY:<date or date-time value>
	attr = t.GetAttribute(MapiPidTagBirthday, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = strconv.FormatInt(int64(attr.GetIntValue()),10)
	}
	propBdayValue := vcard.NewText(attrValue)

	if !propBdayValue.IsEmpty() {
		propBday := vc.CreateProperty("bday")
		propBday.AddValue(propBdayValue)

		vc.AddProperty(propBday)
	}

	// ADR - WORK
	adrWork := vcard.NewAddress()

	// PO BOX
	attr = t.GetAttribute(MapiPidLidWorkAddressPostOfficeBox, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrWork.Pobox = attrValue

	// street
	attr = t.GetAttribute(MapiPidLidWorkAddressStreet, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrWork.Street = attrValue
	//fmt.Println("PidLidWorkAddressStreet: ", attrValue)

	attr = t.GetAttribute(MapiPidLidWorkAddressCity, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrWork.Locality = attrValue
	//fmt.Println("PidLidWorkAddressCity: ", attrValue)

	attr = t.GetAttribute(MapiPidLidWorkAddressState, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrWork.Region = attrValue
	//fmt.Println("PidLidWorkAddressState: ", attrValue)

	attr = t.GetAttribute(MapiPidLidWorkAddressPostalCode, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrWork.Region = attrValue
	//fmt.Println("PidLidWorkAddressPostalCode: ", attrValue)

	attr = t.GetAttribute(MapiPidLidWorkAddressCountry, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrWork.Country = attrValue

	if !adrWork.IsEmpty() {
		propAdrWork := vc.CreateProperty("adr")
		propAdrWork.AddValue(adrWork)

		propAdrWorkParamType := vcard.NewParameter("type")
		propAdrWorkParamType.AddValue("work")
		propAdrWork.AddParameter(propAdrWorkParamType)

		vc.AddProperty(propAdrWork)
	}



	// ADR - home
	adrHome := vcard.NewAddress()

	attr = t.GetAttribute(MapiPidTagHomeAddressPostOfficeBox, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.Pobox = attrValue
	//fmt.Println("PidTagHomeAddressPostOfficeBox: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHomeAddressStreet, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.Street = attrValue
	//fmt.Println("PidTagHomeAddressStreet: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHomeAddressCity, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.Locality = attrValue
	//fmt.Println("PidTagHomeAddressCity: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHomeAddressStateOrProvince, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.Region = attrValue
	//fmt.Println("PidTagHomeAddressStateOrProvince: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHomeAddressPostalCode, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.PostalCode = attrValue
	//fmt.Println("PidTagHomeAddressPostalCode: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHomeAddressCountry, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()

	}
	adrHome.Country = attrValue
	//	fmt.Println("PidTagHomeAddressCountry: ", attrValue)

	if !adrHome.IsEmpty() {
		propAdrHome := vc.CreateProperty("adr")
		propAdrHome.AddValue(adrHome)

		propAdrHomeParamType := vcard.NewParameter("type")
		propAdrHomeParamType.AddValue("home")
		propAdrHome.AddParameter(propAdrHomeParamType)

		vc.AddProperty(propAdrHome)
	}

	// ADR - OTHER

	adrOther := vcard.NewAddress()

	attr = t.GetAttribute(MapiPidTagOtherAddressPostOfficeBox, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.Pobox = attrValue
	//fmt.Println("PidTagOtherAddressPostOfficeBox: ", attrValue)

	attr = t.GetAttribute(MapiPidTagOtherAddressStreet, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.Street = attrValue
	//fmt.Println("PidTagOtherAddressStreet: ", attrValue)

	attr = t.GetAttribute(MapiPidTagOtherAddressCity, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.Locality = attrValue
	//fmt.Println("PidTagOtherAddressCity: ", attrValue)

	attr = t.GetAttribute(MapiPidTagOtherAddressStateOrProvince, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.Region = attrValue
	//fmt.Println("PidTagOtherAddressStateOrProvince: ", attrValue)

	attr = t.GetAttribute(MapiPidTagOtherAddressPostalCode, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.PostalCode = attrValue
	//fmt.Println("PidTagOtherAddressPostalCode: ", attrValue)

	attr = t.GetAttribute(MapiPidTagOtherAddressCountry, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	adrOther.Country = attrValue
	//fmt.Println("PidTagOtherAddressCountry: ", attrValue)

	if !adrOther.IsEmpty() {
		// The "intl", "dom", "postal", and "parcel" TYPE parameter values for the ADR property have been removed (v3.0).
		propAdrOther := vc.CreateProperty("adr")
		propAdrOther.AddValue(adrOther)

		propAdrOtherParamType := vcard.NewParameter("type")
		propAdrOtherParamType.AddValue("postal")
		propAdrOther.AddParameter(propAdrOtherParamType)

		vc.AddProperty(propAdrOther)
	}


	/**
	: If the TYPE parameter contains "pref", the PidLidPostalAddressId property ([MSOXOCNTC] section 2.2.1.3.9) is set to indicate that that address is the contact's mailing address, and
the Mailing Address properties of the Contact object are set as specified in [MS-OXOCNTC] section
2.2.1.3.9.
When exporting: The address that is selected as the mailing address by the PidLidPostalAddressId
property gets the value "pref" included in its TYPE parameter.
	*/

	attr = t.GetAttribute(MapiPidLidPostalAddressId, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	//fmt.Println("PidLidPostalAddressId: ", attrValue)


	//TEL -> : TEL; TYPE=[Type]:[Phone Number]

	// TEL - home
	attr = t.GetAttribute(MapiPidTagHomeTelephoneNumber, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pTelHomeValue := vcard.NewText(attrValue)

	if !pTelHomeValue.IsEmpty() {
		propTelHome := vc.CreateProperty("tel")
		propTelHome.AddValue(pTelHomeValue)

		propTelParamType := vcard.NewParameter("type")
		propTelParamType.AddValue("home")
		propTelHome.AddParameter(propTelParamType)

		vc.AddProperty(propTelHome)
	}
	//fmt.Println("PidTagHomeTelephoneNumber: ", attrValue)

	attr = t.GetAttribute(MapiPidTagHome2TelephoneNumber, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pTelHome2Value := vcard.NewText(attrValue)

	if !pTelHome2Value.IsEmpty() {
		propTelHome2 := vc.CreateProperty("tel")
		propTelHome2.AddValue(pTelHome2Value)

		propTelHome2ParamType := vcard.NewParameter("type")
		propTelHome2ParamType.AddValue("home")
		propTelHome2.AddParameter(propTelHome2ParamType)

		vc.AddProperty(propTelHome2)
	}

	//fmt.Println("PidTagHome2TelephoneNumber: ", attrValue)

	/**
	The first TEL type with the TYPE parameter set to "msg", "voice", "video", "bbs", or "modem"
is imported to PidTagOtherTelephoneNumber ([MS-OXOCNTC] section 2.2.1.4.10).
Additional TEL types with the TYPE parameter set to "msg", "voice", "video", "bbs", or
"modem" are ignored.
	*/
	attr = t.GetAttribute(MapiPidTagOtherTelephoneNumber, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pTelMsgValue := vcard.NewText(attrValue)

	if !pTelMsgValue.IsEmpty() {
		propTelMsg := vc.CreateProperty("tel")
		propTelMsg.AddValue(pTelMsgValue)

		propTelMsgParamType := vcard.NewParameter("type")
		propTelMsgParamType.AddValue("msg")
		propTelMsg.AddParameter(propTelMsgParamType)

		vc.AddProperty(propTelMsg)
	}
	//fmt.Println("PidTagOtherTelephoneNumber: ", attrValue)


	// TEL - WORK
	attr = t.GetAttribute(MapiPidTagBusinessTelephoneNumber, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pTelWorkValue := vcard.NewText(attrValue)

	if !pTelWorkValue.IsEmpty() {
		propTelWork := vc.CreateProperty("tel")
		propTelWork.AddValue(pTelWorkValue)

		propTelWorkParamType := vcard.NewParameter("type")
		propTelWorkParamType.AddValue("work")
		propTelWork.AddParameter(propTelWorkParamType)

		vc.AddProperty(propTelWork)
	}
	//fmt.Println("PidTagBusinessTelephoneNumber: ", attrValue)

	attr = t.GetAttribute(MapiPidTagBusiness2TelephoneNumber, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pTelWork2Value := vcard.NewText(attrValue)

	if !pTelWork2Value.IsEmpty() {
		propTelWork2 := vc.CreateProperty("tel")
		propTelWork2.AddValue(pTelWork2Value)

		propTelWork2ParamType := vcard.NewParameter("type")
		propTelWork2ParamType.AddValue("work")
		propTelWork2.AddParameter(propTelWork2ParamType)

		vc.AddProperty(propTelWork2)
	}
	//fmt.Println("PidTagBusiness2TelephoneNumber: ", attrValue)


	// EMAIL;TYPE=[Type]:[Email]
	attr = t.GetAttribute(MapiPidLidEmail1EmailAddress, "mapi")
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pEmailValue := vcard.NewText(attrValue)
	if !pEmailValue.IsEmpty() {
		propEmail := vc.CreateProperty("email")
		propEmail.AddValue(pEmailValue)
		vc.AddProperty(propEmail)
	}
	//fmt.Println("PidLidEmail1EmailAddress: ", attrValue)

	attr = t.GetAttribute(MapiPidLidEmail2EmailAddress, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	pEmail2Value := vcard.NewText(attrValue)
	if !pEmail2Value.IsEmpty() {
		propEmail2 := vc.CreateProperty("email")
		propEmail2.AddValue(pEmail2Value)
		vc.AddProperty(propEmail2)
	}
	//fmt.Println("PidLidEmail2EmailAddress: ", attrValue)

	attr = t.GetAttribute(MapiPidLidEmail3EmailAddress, "mapi");  // (0x80A3 in docs, but 0x803A in other implementations and samples)
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	if attrValue == "" {
		attr = t.GetAttribute(MapiPidLidEmail3EmailAddress_1, "mapi");
		attrValue = ""
		if attr != nil {
			attrValue = attr.GetStringValue()
		}
	}
	pEmail3Value := vcard.NewText(attrValue)
	if !pEmail3Value.IsEmpty() {
		propEmail3 := vc.CreateProperty("email")
		propEmail3.AddValue(pEmail3Value)
		vc.AddProperty(propEmail3)
	}
	//fmt.Println("PidLidEmail3EmailAddress: ", attrValue)

	//PidTagProfession - ROLE
	attr = t.GetAttribute(MapiPidTagProfession, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	roleValue := vcard.NewText(attrValue)
	if !roleValue.IsEmpty() {
		roleProp := vc.CreateProperty("role")
		roleProp.AddValue(roleValue)
		vc.AddProperty(roleProp)
	}

	//PidTagProfession - ORG

	 // ORG - company (PidTagCompanyName)
	 orgValue := vcard.NewOrganization("", []string{})

	attr = t.GetAttribute(MapiPidTagCompanyName, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	orgValue.Company = attrValue


	 // ORG - department (PidTagDepartmentName)
	attr = t.GetAttribute(MapiPidTagDepartmentName, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}
	if attrValue != "" {
		orgValue.Departments = append(orgValue.Departments, attrValue)
	}

	if !orgValue.IsEmpty() {
		orgProp := vc.CreateProperty("org")
		orgProp.AddValue(orgValue)
		vc.AddProperty(orgProp)
	}


	 // CATEGORIES -PidLidCategories
	 categProp := vcard.NewProperty("categories")
	 attr = t.GetAttribute(0x00009000, "mapi");
	 attrValue = ""
	 if attr != nil {
		for _, attrValue = range attr.GetStringValueArray() {
			if attrValue != "" {
				categProp.AddValue(vcard.NewText(attrValue))
			}
		 }
	 }

	 if len(categProp.GetValue()) > 0  {
		vc.AddProperty(categProp)
	 }

	 // NOTE - PidTagBody
	 attr = t.GetAttribute(MapiPidTagBody, "mapi")
	 attrValue = ""
	 if attr != nil {
		 attrValue = attr.GetStringValue()
	 }

	 noteValue := vcard.NewText(attrValue)
	 if !noteValue.IsEmpty() {
		 noteProp := vc.CreateProperty("note")
		 noteProp.AddValue(noteValue)
		 vc.AddProperty(noteProp)
	 }

	 // REV - PidTagLastModificationTime
	 attr = t.GetAttribute(AttDateModified, "mapped")
	 attrValue = ""
	 if (attr != nil) {
		attrValue = strconv.FormatInt(int64(attr.GetIntValue()),10)
	 }


	 revValue := vcard.NewText(attrValue)
	 if !revValue.IsEmpty() {
		 revProp := vc.CreateProperty("rev")
		 revProp.AddValue(revValue)
		 vc.AddProperty(revProp)
	 }

	 //URL - PidTagPersonalHomePage
	 attr = t.GetAttribute(MapiPidTagPersonalHomePage, "mapi")
	 attrValue = ""
	 if attr != nil {
		 attrValue = attr.GetStringValue()
	 }

	 urlHomeValue := vcard.NewText(attrValue)
	 if !urlHomeValue.IsEmpty() {
		 urlHomeProp := vc.CreateProperty("url")
		 urlHomeProp.AddValue(urlHomeValue)

		 urlHomeParamType := vcard.NewParameter("type")
		 urlHomeParamType.AddValue("home")
		 urlHomeProp.AddParameter(urlHomeParamType)

		 vc.AddProperty(urlHomeProp)
	 }

	 //URL -
	 attr = t.GetAttribute(MapiPidTagBusinessHomePage, "mapi")
	 attrValue = ""
	 if attr != nil {
		 attrValue = attr.GetStringValue()
	 }

	 urlWorkValue := vcard.NewText(attrValue)
	 if !urlWorkValue.IsEmpty() {
		 urlWorkProp := vc.CreateProperty("url")
		 urlWorkProp.AddValue(urlWorkValue)

		 urlWorkParamType := vcard.NewParameter("type")
		 urlWorkParamType.AddValue("work")
		 urlWorkProp.AddParameter(urlWorkParamType)

		 vc.AddProperty(urlWorkProp)
	 }


	 // CLASS - PidTagSensitivity -> removed in vcard v4.0

		attr = t.GetAttribute(MapiPidTagSensitivity, "mapi");
		attrValue = ""
		if attr != nil {
			switch attr.GetIntValue() {
				case 0:
					attrValue = "PUBLIC"
				case 2:
					attrValue = "PRIVATE"
				case 3:
					attrValue = "CONFIDENTIAL"
			}
		}

		classValue := vcard.NewText(attrValue)
		if !classValue.IsEmpty() {
			classProp := vc.CreateProperty("class")
			if classProp != nil {
				classProp.AddValue(classValue)
				vc.AddProperty(classProp)
			}
		}

	 // KEY - PidTagUserX509Certificate
	 /*
	 attr = t.GetMapiAttribute(0x3A70);
	 attrValue = ""
	 if attr != nil {
		 //@TODO: extract correct certificat value (PidTagUserX509Certificate [MS-OXOABK] - remove first 4 bytes - 2 bytes for tag & 2 byte for length)
		 attrValue = strings.TrimRight(string(attr.Data), "\x00")
	 }

	 keyValue := vcard.NewText()
	 keyValue.SetValue(attrValue)
	 if !keyValue.IsEmpty() {
		keyProp := vcard.NewProperty("key")
		keyProp.AddValue(keyValue)

		keyParamType := vcard.NewParameter("encoding")
		keyParamType.AddValue("b")
		keyProp.AddParameter(keyParamType)

		 vc.AddProperty(keyProp)
	 }*/

	 //X-MS-OL-DESIGN
	 attr = t.GetAttribute(MapiPidLidBusinessCardDisplayDefinition, "mapi")
	 attrValue = ""
	 if attr != nil {
		 // @ToDO - decode binary [MS-OXOCNTC] 2.2.1.7.1
	 }

	 designValue := vcard.NewText(attrValue)
	 if !designValue.IsEmpty() {
		designProp := vc.CreateProperty("X-MS-OL-DESIGN")
		designProp.AddValue(designValue)
		vc.AddProperty(designProp)
	 }


	//FBURL - available only on 4.0
	attr = t.GetAttribute(MapiPidLidFreeBusyLocation, "mapi");
	attrValue = ""
	if attr != nil {
		attrValue = attr.GetStringValue()
	}

	fburlValue := vcard.NewText(attrValue)
	if !fburlValue.IsEmpty() {
		fburlProp := vc.CreateProperty("fburl")
		if fburlProp != nil {
			fburlProp.AddValue(fburlValue)
			vc.AddProperty(fburlProp)
		}
	}
}
