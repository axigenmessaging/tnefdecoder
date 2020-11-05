package tnefdecoder

import (
	"errors"
)

var ErrNoMarker = errors.New("file did not begin with a TNEF marker")

const (
	TnefSignature 		= 0x223e9f78
	AttrLevelMessage    = 0x01
	AttrLevelAttachment = 0x02
)

/**
  *	 Message-level TNEF mapped attribute (these are mapi properties mapped into tnef properties). SHOULD all be at attrLevelMessage;
  *  the other MAPI properties are encoded into AttMsgProps
  */
const (
	// OEM Codepage. See attOemCodepage handling in section 5 and Appendix
	AttOEMCodepage =  0x00069007

	//TNEF Version
	AttTnefVersion =  0x00089006

	//PidTagMessageClass property ([MS-OXOMSG] section 2.2.2.1)  %x08.80.07.00
	AttMessageClass =  0x00078008

	// PidTagSender_XXX properties, as specified in section 2.1.3.1.1
	AttFrom = 0x00008000

	// PidTagSubject property ([MS-OXCMSG] section 2.2.1.46)   %x04.80.01.00
	AttSubject = 0x00018004

	// PidTagClientSubmitTime property ([MS-OXOMSG] section 2.2.3.11) %x05.80.03.00
	AttDateSent = 0x00038005

	// PidTagMessageDeliveryTime property ([MS-OXOMSG] section 2.2.3.9) %x06.80.03.00
	AttDateRecd = 0x00038006

	// PidTagMessageFlags property ([MS-OXCMSG] section 2.2.1.6) %x07.80.06.00
	AttMessageStatus = 0x00068007

	// PidTagSearchKey property ([MS-OXCPRPT] section 2.2.1.9) %x09.80.01.00
	AttMessageID = 0x00018009

	// MAP to PidTagBody property ([MS-OXCMSG] section 2.2.1.56.1) %x0C.80.02.00
	AttBody = 0x0002800C

	// PidTagImportance property ([MS-OXCMSG] section 2.2.1.11) %x0D.80.04.00
	AttPriority = 0x0004800D

	// PidTagLastModificationTime property ([MS-OXCMSG] section 2.2.2.2) for the message %x20.80.03.00
	AttDateModified = 0x00038020

	// Message Property Encapsulation %x03.90.06.00
	AttMsgProps = 0x00069003

	// PidTagMessageRecipients propety ([MS-OXCMSG] section 2.2.1.47).
	// For more details, see section 2.1.3.5.2.2. %x04.90.06.00
	AttRecipTable = 0x00069004

	// PidTagOriginalMessageClass property ([MS-OXPROPS] section 2.824) %x00.06.07.00
	AttOriginalMessageClass = 0x00070600

	// PidTagReceivedRepresenting_XXX or PidTagSentRepresenting_XXX properties, as specified in section 2.1.3.1.1 %x00.00.06.00
	AttOwner = 0x00060000

	// PidTagSentRepresenting_XXX properties, as specified in section 2.1.3.1.1 %x01.00.06.00
	AttSentFor = 0x00060001

	// PidTagReceivedRepresentingEntryId property ([MS-OXOMSG] section 2.2.1.25) %x02.00.06.00
	AttDelegate = 0x00060002

	// PidTagStartDate property ([MS-OXOCAL] section 2.2.1.30) %x06.00.03.00
	AttDateStart = 0x00030006

	// PidTagEndDate property ([MS-OXOCAL] section 2.2.1.31) %x07.00.03.00
	AttDateEnd = 0x00030007

	// PidTagOwnerAppointmentId property ([MS-OXOCAL] section 2.2.1.29) %x08.00.05.00
	AttAidOwner = 0x00050008

	// PidTagResponseRequested property ([MS-OXOCAL] section 2.2.1.36) %x09.00.04.00
	AttRequestRes = 0x00040009
)

/*
 *  Attachment mapped attributes
 * Attachment-level attributes. All MUST be at attrLevelAttachment. (mapped mapi attributes to Attachment)
 * idAttachAttr = idAttachData / idAttachTitle / idAttachMetaFile /
 *	IdAttachCreateDate / idAttachModifyDate / idAttachTransportFilename
 */
const (

	// PidTagAttachDataBinary property ([MS-OXCMSG] section 2.2.2.7) %x0F.80.06.00
	AttAttachData = 0x0006800F

	// mapped attachment attr of  mapi attribute PidTagAttachLongFilename if available, otherwise PidTagAttachFilename property ([MS-OXCMSG] section 2.2.2.11) %x10.80.01.00
	AttAttachTitle = 0x00018010

	// PidTagAttachRendering property ([MS-OXCMSG] section 2.2.2.17) %x11.80.06.00
	AttAttachMetaFile = 0x00068011

	// PidTagCreationTime property ([MS-OXCMSG] section 2.2.2.3) %x12.80.03.00
	AttAttachCreateDate = 0x00038012

	// PidTagLastModificationTime property ([MS-OXCMSG] section 2.2.2.2)
	// for the attachment. %x13.80.03.00
	AttAttachModifyDate = 0x00038013

	// PidTagAttachTransportName property ([MS-OXCMSG] section 2.2.2.19) %x01.90.06.00
	AttAttachTransportFilename = 0x00069001

	// Attachment RendData, as specified in section 2.1.3.3.15 %x02.90.06.00
	AttAttachRendData = 0x00069002

	// Attachment table row %x05.90.06.00
	AttAttachment = 0x00069005
)


/**
 * mapi attributes data types
 */
 const (
	MapiTypeUnspecified = 0x0000 // type unspecified
	MapiTypeNull = 0x0001 // null
	MapiTypeInt16 = 0x0002 // int 16
	MapiTypeMVInt16 = 0x1002 //	array of int16
	MapiTypeInt32 = 0x0003 //int32
	MapiTypeMVInt32 = 0x1003 // array of int32
	MapiTypeFlt32 = 0x0004 // float32
	MapiTypeMVFlt32 = 0x1004 // array of float32
	MapiTypeFlt64 = 0x0005 //float 64
	MapiTypeMVFlt64 = 0x1005 // array of float32
	MapiTypeCurrency = 0x0006 //Signed 64-bit
	MapiTypeMVCurrency = 0x1006 // array of Signed 64-bit
	MapiTypeAppTime = 0x0007 // float64
	MapiTypeMVAppTime = 0x1007 // array of float64
	MapiTypeBoolean = 0x000B // 16-bit Boolean (non-zero = TRUE)
	MapiTypeObject = 0x000D // Embedded object on a property (unicode?)
	MapiTypeInt64 = 0x0014 // 8-byte signed integer = INT64.
	MapiTypeMVInt64 = 0x1014 // array of int64
	MapiTypeString8 = 0x001E //8-bit character string with terminating null character
	MapiTypeMVString8 = 0x101E // array of 8-bit character string with terminating null character.
	MapiTypeUnicode  = 0x001F // TypeUnicode - UTF-16LE or variant character string with terminating 2-byte null character.
	MapiTypeMVUnicode = 0x101F // array of unicode
	MapiTypeSystime = 0x0040 //FILETIME (a PtypTime value, as specified in [MS-OXCDATA] section 2.11.1)
	MapiTypeMVSystime = 0x1040 // array of TypeSystime
	MapiTypeCLSID = 0x0048 //OLE GUID - 16 bytes
	MapiTypeMVCLSID = 0x1048 // array of TypeCLSID
	MapiTypeBinary = 0x0102 // binary
	MapiTypeMVBinary = 0x1102 // array of binary
)


/**
 * MAPI Attributes
 */

 // attachment mapi attributes (extacted from attAttachment)
 const (
	MapiPidTagAttachDataBinary = 0x3701 // PidTagAttachDataBinary - Contains the contents of the file to be attached.
	MapiPidTagAttachSize = 0xe20 //Type: 0x0003 -> PidTagAttachSize | value: 3285 (bytes)
	MapiPidTagDisplayName	= 0x3001 // TAG Type: 30 (0x001e) -> PidTagDisplayName (type: 0x001f) | value: image001.jpg (same as PidTagAttachLongFilename) + display name to vcard
	MapiPidTagAttachEncoding = 0x3702 //TAG Type: 258 (0x0102) -> PidTagAttachEncoding | value: empty!!?? ->  If the attachment is in MacBinary format, this property is set to "{0x2A,86,48,86,F7,14,03,0B,01}"; otherwise, it is unset.
	MapiPidTagAttachExtension = 0x3703 // TAG Type: 30 (0x001e) -> PidTagAttachExtension (type: 0x001e) | value: .jpg
	MapiPidTagAttachMethod = 0x3705 //TAG Type: 3 (0x0003) -> PidTagAttachMethod | value: 1
	MapiPidTagAttachLongFilename = 0x3707 // TAG Type: 30 (0x1e) -> PidTagAttachLongFilename (0x001F) | value: image001.jpg - (string) Contains the full filename and extension of the Attachment object.
	MapiPidTagAttachFilename =  0x3704 // string -contains the 8.3 name of the filename

	MapiPidTagRenderingPosition = 0x370b // TAG Type: 3 (0x0003) -> PidTagRenderingPosition | value: -1 (-1 e de fapt 0xffffff, decoded as signed) ->  0xFFFFFFFF indicates a hidden attachment that is not to be rendered in the main text
	MapiPidTagAttachMimeTag = 0x370e // TAG Type: 30 (0x001e) -> PidTagAttachMimeTag | value: image/jpeg  (string) - Contains a content-type MIME header.
	MapiPidTagAttachFlags = 0x3714 // TAG Type: 3 (0x3) -> PidTagAttachFlags | value: 4 (4 means attRenderedInBody)
	MapiPidTagAttachmentLinkId = 0x7ffa //, TAG Type: 3 (0x3) -> PidTagAttachmentLinkId| value: 0 (must be 0, if is not overwriten)
	MapiPidTagExceptionStartTime = 0x7ffb // TAG Type: 64 (0x0040) ->	PidTagExceptionStartTime|value: 915151392000000000
	MapiPidTagExceptionEndTime =  0x7ffc // TAG Type: 64 (0x40) -> PidTagExceptionEndTime | value: 915151392000000000
	MapiPidTagAttachmentFlags = 0x7ffd //, TAG Type: 3 (0x3) -> PidTagAttachmentFlags | value: 8
	MapiPidTagAttachmentHidden = 0x7ffe // TAG Type: 11 (0xb) -> PidTagAttachmentHidden| value: true
	MapiPidTagAttachmentContactPhoto = 0x7fff // MAPI ATTR ID: 32767 (0x7fff), TAG Type: 11 (0xb) -> PidTagAttachmentContactPhoto | value: false
	MapiPidTagAttachNumber = 0x0e21 //, TAG Type: 3 (0x3) -> PidTagAttachNumber | value: 956325
	MapiPidTagMappingSignature  = 0x0ff8  // TAG Type: 258 (0x0102) -> PidTagMappingSignature | value: 28 78 81 160 198 126 89 69 167 247 18 51 167 63 155 237
	MapiPidTagObjectType =  0xffe //, TAG Type: 3 (0x3) -> PidTagObjectType | value: 7 (7 means Attachment object)
	MapiPidTagStoreSupportMask = 0x340d// TAG Type: 3 (0x3) ->  | value: 245710845 ( Indicates whether string properties within the .msg file are Unicode-encoded.)

	MapiPidTagTnefCorrelationKey = 0x007F // binary - Contains a value that correlates a Transport Neutral Encapsulation Format (TNEF) attachment with a message.
	MapiPidTagAttachContentId = 0x3712 // string - Contains a content identifier unique to the Message object that matches a corresponding "cid:" URI schema reference in the HTML body of the Message object. (ex: image001.jpg@01D49162.DB2DC760)
)



/**
 * MAPI Body Properties
 */

 const (
 	MapiPidTagBody = 0x1000 // string The PidTagBody property ([MS-OXPROPS] section 2.618) contains unformatted text, which is the text/plain MIME format
 	MapiPidTagNativeBody = 0x1016 // int -  PidTagNativeBody property ([MS-OXPROPS] section 2.805) indicates the best available format for storing the message body
 	MapiPidTagBodyHtml = 0x1013 // string - PidTagBodyHtml property ([MS-OXPROPS] section 2.621) contains the HTML body
 	MapiPidTagRtfCompressed = 0x1009 // binary -  PidTagRtfCompressed property ([MS-OXPROPS] section 2.941) contains an RTF body compressed
 	MapiPidTagRtfInSync = 0x0E1F //  PidTagRtfInSync property ([MS-OXPROPS] section 2.942) is set to "TRUE" (0x01) if the RTF body has been synchronized with the contents in the PidTagBody (Indicates whether the PidTagBody property (section 2.618) and the	 PidTagRtfCompressed property (section 2.941) contain the same text (ignoring formatting).)
 	MapiPidTagInternetCodepage = 0x3FDE // int32 The PidTagInternetCodepage property ([MS-OXPROPS] section 2.746) indicates the code page used for the PidTagBody property (section 2.2.1.56.1) or the PidTagBodyHtml property
	MapiPidTagBodyContentId = 0x1015 // PidTagBodyContentId property ([MS-OXPROPS] section 2.619) contains a GUID corresponding	to the current message body. This property corresponds to the Content-ID header.
 )

 /**
  * MAPI contact properties used for vcard
  */
 const (

	MapiPidTagNormalizedSubject = 0x0E1D // PtypString - The PidTagNormalizedSubject property ([MS-OXCMSG] section 2.2.1.10) specifies a combination of	the full name and company name of the contact
	MapiPidTagConversationTopic = 0x0070 //  string - The PidTagConversationTopic property ([MS-OXPROPS] section 2.646) contains an unchanging 	copy of the original subject.<4> The property is set to the same value as the 	PidTagNormalizedSubject property ([MS-OXCMSG] section 2.2.1.10) on an E-mail object when it	is submitted.

	MapiPidTagSurname = 0x3A11 // string
	MapiPidTagGivenName = 0x3A06 // string
	MapiPidTagMiddleName = 0x3A44 // string
	MapiPidTagDisplayNamePrefix = 0x3A45 // string
	MapiPidTagGeneration = 0x3A05 //string
	MapiPidTagNickname = 0x3A4F // string


	MapiPidTagBirthday = 0x3A42 // PtypTime

	MapiPidLidWorkAddressPostOfficeBox = 0x804A // string
	MapiPidLidWorkAddressStreet = 0x8045 //string
	MapiPidLidWorkAddressCity = 0x8046 //string
	MapiPidLidWorkAddressState = 0x8047 //string
	MapiPidLidWorkAddressPostalCode = 0x8048 //string
	MapiPidLidWorkAddressCountry = 0x8049 // string
	MapiPidTagHomeAddressPostOfficeBox = 0x3A5E //string
	MapiPidTagHomeAddressStreet = 0x3A5D // string
	MapiPidTagHomeAddressCity = 0x3A59 //string
	MapiPidTagHomeAddressStateOrProvince = 0x3A5C //string
	MapiPidTagHomeAddressPostalCode = 0x3A5B //string
	MapiPidTagHomeAddressCountry = 0x3A5A //string
	MapiPidTagOtherAddressPostOfficeBox = 0x3A64 //string
	MapiPidTagOtherAddressStreet = 0x3A63 //string
	MapiPidTagOtherAddressCity = 0x3A5F //string
	MapiPidTagOtherAddressStateOrProvince = 0x3A62 //string
	MapiPidTagOtherAddressPostalCode = 0x3A61 // string
	MapiPidTagOtherAddressCountry = 0x3A60 //string
	MapiPidLidPostalAddressId = 0x8022 //string
	MapiPidTagHomeTelephoneNumber = 0x3A09 //string
	MapiPidTagHome2TelephoneNumber = 0x3A2F // string
	MapiPidTagOtherTelephoneNumber = 0x3A1F //string
	MapiPidTagBusinessTelephoneNumber = 0x3A08 //string
	MapiPidTagBusiness2TelephoneNumber = 0x3A1B //string
	MapiPidLidEmail1EmailAddress = 0x8083 // string
	MapiPidLidEmail2EmailAddress = 0x8093 // string
	MapiPidLidEmail3EmailAddress = 0x80A3 // string
	MapiPidLidEmail3EmailAddress_1 = 0x803A // string - equivalend of MapiPidLidEmail3EmailAddress found in other implementations
	MapiPidTagProfession = 0x3A46 //string
	MapiPidTagCompanyName = 0x3A16 //string
	MapiPidTagDepartmentName = 0x3A18 //string
	MapiPidLidCategories = 0x00009000 // string array
	MapiPidTagPersonalHomePage = 0x3A50 //string
	MapiPidTagBusinessHomePage = 0x3A51 //string
	MapiPidTagSensitivity = 0x0036 // int32
	MapiPidLidBusinessCardDisplayDefinition = 0x00008040 // binary -> decoding info [MS-OXOCNTC] 2.2.1.7.1
	MapiPidLidFreeBusyLocation = 0x000080D8 //string
	MapiPidLidHasPicture = 0x00008015 // bool - The PidLidHasPicture property ([MS-OXPROPS] section 2.144) indicates whether a contact photo attachment, specified in section 2.2.1.8.3, exists. If this property is set to nonzero (TRUE), then the contact photo attachment exists and the client uses it as the contact photo.
	MapiPidTagLastModificationTime = 0x3008 //PT_SYSTIME -> int64

 )

