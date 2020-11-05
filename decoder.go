// Package tnef extracts the body and attachments from Microsoft TNEF files.
package tnefdecoder

import (
	"errors"
	"bytes"
	"io/ioutil"
	"vcard"
	"fmt"
)


func NewDecoder() TnefDecoder {
	d := TnefDecoder{}
	d.VcardVersion = "3.0"
	d.leDecoder = new(LittleEndianDecoder)

	return d
}


type TnefDecoder struct {
	VcardVersion string
	leDecoder *LittleEndianDecoder
}



/** DecodeFile is a utility function that reads the file into memory
 *  before calling the normal Decode function on the data.
 */
func (d *TnefDecoder) DecodeFile(path string) (*TnefObject, error) {
	data, err := ioutil.ReadFile(path)
	if err != nil {
		return nil, err
	}

	return d.Decode(data)
}

/**
 * decode TNEF bytes
 */
 func (d *TnefDecoder) Decode(data []byte) (*TnefObject, error) {
	var tAttachment *Attachment

	tObj := &TnefObject{}

	offset := 0
	// TNEFSignature signals that the file did not start with the fixed TNEF marker,
	// meaning it's not in the TNEF file format we recognize (e.g. it just has the
	// .tnef extension, or a wrong MIME type).
	if d.leDecoder.Int(data[0:4]) != TnefSignature {
		return nil, errors.New("file did not begin with a TNEF marker")
	}
	offset += 4

	dataLength := len(data)

	// LegacyKey UINT16 - 2bytes - ignored
	offset += 2

	// TNEFVersion
	for {
		attr, noByteRead := d.DecodeAttributeStructure(data[offset:])

		//fmt.Printf("\r\n ATTR: type: %s Level: %v ID: 0x%X | %v, Length: %v  => %s",  attr.Type,attr.Level, attr.Id, attr.Id, len(attr.Data), attr.Data)


		if  (attr.Level == AttrLevelMessage) {
			// message attributes

			// reset the attachment
			tAttachment = nil

			// message attributes
			tObj.Attributes = append(tObj.Attributes, attr)

			/**
			 * some attributes require special decoding
			 */
			switch (attr.Id) {
				case AttMessageID:
					// is tnef attribute (mapped MAPI)
				case AttMsgProps:
					// this attribute contains mapi attributes  - we must extract properties
					attrList, err := d.DecodeMapiProperties(attr.Data)
					if err == nil && len(attrList) > 0 {
						tObj.Attributes = append(tObj.Attributes, attrList...)
					}
				case AttRecipTable:
					// recipient table
				default:
					tObj.Attributes = append(tObj.Attributes, attr)
			}

		} else if  (attr.Level == AttrLevelAttachment) {
			// attachment attributes

			/**
			 * Each set of attachment attributes MUST begin with the attAttachRendData attribute, followed by any
			 * other attributes; attachment properties encoded in the attAttachment attribute SHOULD be last
			 * apply special decoding on specific attributes
			 */
			switch (attr.Id) {
				case AttAttachRendData:
					// start a new attachemnt
					/**
					attAttachRendData = AttachType AttachPosition RenderWidth RenderHeight DataFlags
					AttachType = AttachTypeFile / AttachTypeOle
					AttachTypeFile=%x01.00
					AttachTypeOle=%x02.00
					AttachPosition= INT32
					RenderWidth=INT16
					RenderHeight=INT16
					DataFlags = FileDataDefault / FileDataMacBinary
					FileDataDefault= %x00.00.00.00
					FileDataMacBinary=%x01.00.00.00
					*/
					tAttachment = NewAttachment()
					tObj.Attachments = append(tObj.Attachments, tAttachment)

					// it's not required, the property seems to be a summary of some other MAPI Attributes; we should find all info in decoded mapi attributes
					// tAttachment.DecodedAttRendData = d.DecodeAttachmentRendData(attr.Data)
				case AttAttachData:
					// it's the body of the attachment
					tAttachment.SetData(attr.Data)
				case AttAttachment:
					// attchment table row -> decode Mapi Attributes
					attrList, err := d.DecodeMapiProperties(attr.Data)
					if err == nil && len(attrList) > 0 {
						tAttachment.Attributes = append(tAttachment.Attributes, attrList...)
					}

					/**
					If the PidTagAttachMethod property ([MS-OXCMSG] section 2.2.2.9) of the original attachment  contains the value 0x0005 (ATTACH_EMBEDDED_MSG) or the value 0x0006 (ATTACH_OLE), then the TNEF Reader SHOULD ignore the attAttachData attribute, as specified in section 2.1.3.3.11.
					*/

					pidTagAttachMethodAttr := tAttachment.GetAttribute(MapiPidTagAttachMethod, "mapi")

					if pidTagAttachMethodAttr != nil && (pidTagAttachMethodAttr.GetIntValue() == 5 || pidTagAttachMethodAttr.GetIntValue() == 6) {

						// decode VCARD
						binaryDataAttr := tAttachment.GetAttribute(MapiPidTagAttachDataBinary, "mapi")
						//objectPrefix := []byte{"\x07", "\x03", "\x02", "\x00", "\x00", "\x00", "\x00", "\x00", "\xC0", "\x00", "\x00", "\x00", "\x00", "\x00", "\x00", "\x46"}

						if binaryDataAttr != nil && binaryDataAttr.DataType == MapiTypeObject {
							objectPrefix := []byte{7, 3, 2, 0, 0, 0, 0, 0, 192, 0, 0, 0, 0, 0, 0, 70}
							streamValue := binaryDataAttr.GetObjectValue()
							if (bytes.HasPrefix(streamValue, objectPrefix)) {
								// the binary data of the attachment is a tnef object
								streamValue = bytes.TrimPrefix(streamValue, objectPrefix)

								attTnefObj, errD := d.Decode(streamValue)

								if (errD == nil && attTnefObj != nil && attTnefObj.GetMessageClass() == "IPM.Contact") {
									// the attachment is a vcard.vcf
									var vcBuilder vcard.IVCard

									switch (d.VcardVersion) {
										case "3.0":
											vcBuilder = vcard.NewVCardV3()
										default:
											vcBuilder = vcard.NewVCardV3()

									}
									ExtractVCard(attTnefObj, vcBuilder)
									tAttachment.SetData([]byte(vcBuilder.Build()))

									vcardFilename := "vcard.vcf"
									fnArr := vcBuilder.GetProperty("fn")
									if len(fnArr) > 0 {
										fnValue := fnArr[0].GetFirstValue()
										if fnValue != nil && fnValue.GetValue() != "" {
											vcardFilename = fnValue.GetValue() + ".vcf"
										}
									}
									tAttachment.SetFilename(vcardFilename)
								}
							}
						}
					}

					//attachMethodAttr := tAttachment.GetAttribute(MapiPidTagAttachMethod, "mapi")
					//fmt.Printf("\r\nMetoda atasament MAPI ID: %#x , Bytes: %v => Val: %v ", MapiPidTagAttachMethod, hex.Dump(attachMethodAttr.Data), attachMethodAttr.GetIntValue())

				default:
					//fmt.Println("Atasamet: ", tAttachment)
					tAttachment.Attributes = append(tAttachment.Attributes, attr)
			}
		}

		offset += noByteRead
		if offset >= dataLength {
			break
		}
	}

	// check if we the TNEF has RTF
	tObj.DecodeRtf()

/*
	fmt.Println("\r\n------------------ START TNEF -----------------------")
	for _, a := range tObj.Attachments {
		fmt.Printf("\r\n\r\nFilename: %v\r\n", a.GetFilename())
		attr := a.GetAttribute(MapiPidTagAttachMethod, "mapi")
		if attr != nil {
			fmt.Printf("\r\nPidTagAttachMethod: %v", attr.GetIntValue())
		}
		attr = a.GetAttribute(MapiPidTagAttachContentId, "mapi")
		if attr != nil {
			fmt.Printf("\r\nPidTagAttachContentId: %v", attr.GetStringValue())
		}

		attr = a.GetAttribute(MapiPidTagAttachExtension, "mapi")
		if attr != nil {
			fmt.Printf("\r\nPidTagAttachExtension: %v", attr.GetStringValue())
		}

		attr = a.GetAttribute(MapiPidLidHasPicture, "mapi");
		if attr != nil {
			fmt.Println("HAS PHOTO: ",attr.GetBoolValue())
		}

		attr = a.GetAttribute(MapiPidTagAttachmentContactPhoto, "mapi");
		if attr != nil {
			fmt.Println("PHOTO: ", attr.GetBoolValue())
		}
	}
	fmt.Println("--------------- END TNEF ----------------------------------")
	*/



	return tObj, nil
}

/**
 *  return the attribute and the total of bytes read used to create attribute
 *  the function decodes the pattern:
 *		levelMessage idAttribute Length Data Checksum
 *
 *  ex:
 *  MessageAttribute = attrLevelMessage idMessageAttr Length Data Checksum
 *  MessageProps = attrLevelMessage idMsgProps Length Data Checksum
 *  AttachRendData = attrLevelAttachment idAttachRendData Length Data Checksum
 *  AttachAttribute = attrLevelAttachment idAttachAttr Length Data Checksum
 *  AttachProps = attrLevelAttachment idAttachment Length Data Checksum
 *
 */

 func (d *TnefDecoder) DecodeAttributeStructure(data []byte) (*Attribute,  int) {

	attr := &Attribute{}
	attr.Type = "mapped"

	offset := 0

	// Level: 1 byte
	attr.Level = d.leDecoder.Int(data[offset : offset+1])
	offset++

	// read attribute ID
	attr.Id = d.leDecoder.Int(data[offset : offset+4])
	offset += 4

	// read attribute data length
	attLength := d.leDecoder.Int(data[offset : offset+4])
	offset += 4

	// read attribute Data (value) - we do not know the type of the value
	attr.Data = data[offset : offset+attLength]
	offset += attLength

	// read attr checksum -> UINT16
	//attr.Checksum = int()
	offset += 2

	return attr, offset
}


/**
 *  extract MAPI attributes from  attMsgProps or attAttachment attributes
 *  the value extracted is []bytes
 *	the value must be extracted with specialized functions for based on MAPI attribute type (int, string, array of type, object, etc)
 *
 *  MsgPropertyList = MsgPropertyCount *MsgPropertyValue
 *  MsgPropertyCount = UINT32
 *  MsgPropertyValue = MsgPropertyTag MsgPropertyData
 *
 * @param  {[type]} data []byte)       (MsgPropertyList [description]
 * @return {[type]}      [description]
 */
 func (d *TnefDecoder) DecodeMapiProperties(data []byte) ([]*Attribute, error) {

	dataLength := len(data)

	if dataLength < 4 {
		return nil, fmt.Errorf("decodeMsgPropertyList: data too short")
	}

	offset := 0

	// no of properties encoded
	noOfAttributes := int(d.leDecoder.Uint32(data[offset:offset+4]))

	list := make([]*Attribute, noOfAttributes)

	offset += 4

	//MsgPropertyValue = MsgPropertyTag MsgPropertyData

	//fmt.Printf("\r\nData MAPI Length: %v")

	for aidx:=0; aidx < noOfAttributes; aidx++ {
		attrDataBuf := bytes.NewBuffer([]byte{})

		attr := &Attribute{}
		attr.Type = "mapi"

		if offset >= dataLength {
			return nil, fmt.Errorf("offset is too large: %d", offset)
		}
		/* MsgPropertyTag = MsgPropertyType MsgPropertyId [NamedPropSpec] */

		// MAPI property value type
		attr.DataType = int(d.leDecoder.Uint16(data[offset:offset+2])) // 2 bytes
		offset += 2

		// MAPI property ID
		attr.Id = int(d.leDecoder.Uint16(data[offset:offset+2])) // 2 bytes
		offset += 2

		if attr.Id >= 0x8000 {
			// has  NamedPropSpec; NamedPropSpec = PropNameSpace PropIDType PropMap
			attr.GUID = d.leDecoder.String(data[offset:offset + 16])
			offset += 16

			attr.PropMapValueType = int(d.leDecoder.Uint32(data[offset:offset+4]))
			offset += 4

			if attr.PropMapValueType == 0x00000000 {
				// should be an uint32 value
				attr.PropMapValue = int(d.leDecoder.Uint32(data[offset:offset+4]))
				offset += 4
			} else {
				// is string
				// propIDType == 0x01000000	=> is PropMap is string (PropMapString)
				// PropMapString = UINT32 *UINT16 %x00.00 [PropMapPad]
				readLength := int(d.leDecoder.Uint32(data[offset:offset+4])) // the length includes the padding
				offset+=4

				attr.PropMapValue = d.leDecoder.Utf16(data[offset:offset+readLength])
				//fmt.Printf("\r\n Custom: Len: %v Value: %v", readLength, attr.PropMapValue)
				//fmt.Printf("\r\nHEX1: %v \r\nHex2: %v", hex.Dump(data[offset:offset + readLength]), hex.Dump(data[offset + readLength: offset + readLength+8]))
				offset += readLength

				// be sure valueLength is; valueLength should be equal with bytesRead + padd
				if padd := 4 - (readLength % 4); padd < 4 {
					offset += padd
				}
			}
		}

		valueBytesLength, isMultiValue := GetTypeSize(attr.DataType)

		//fmt.Printf("\r\nDataType: %#x LengthBytes: %v IsMulti: %v Offset: %v", attr.DataType, valueBytesLength, isMultiValue, offset)

		countAttrValues := 1
		if isMultiValue {
			countAttrValues = int(d.leDecoder.Uint32(data[offset:offset + 4]))
			//fmt.Printf("\r\nValue Count: %#x ", countAttrValues)
			attrDataBuf.Write(data[offset:offset+4])
			offset += 4
		}

		for i := 0; i < countAttrValues; i++ {
			if (valueBytesLength == -1) {
				// variable content
				valueBytesLength = int(d.leDecoder.Uint32(data[offset:offset + 4]))
				attrDataBuf.Write(data[offset:offset+4])
				offset += 4
			}
			if padd := 4 - (valueBytesLength % 4); padd < 4 {
				valueBytesLength += padd
			}

			if offset + valueBytesLength > dataLength {
				return nil, fmt.Errorf("offset is too large when extracting value : %d", offset + valueBytesLength)
			}
			attrDataBuf.Write(data[offset:offset+valueBytesLength])
			offset += valueBytesLength
		}



		/**
		 * the attribute Data contains the logic for no of values and value size (for variable content) to be decoded when the value is need it
		 */
		attr.Data = attrDataBuf.Bytes()


		//fmt.Printf("\r\nDecode -> MAPI Type: %#x, ID: %#x Data: %v Offset: %v", attr.DataType, attr.Id, hex.Dump(attr.Data), offset)

		list[aidx] = attr
	}

	return list, nil
}



/**
 * attAttachRendData = AttachType AttachPosition RenderWidth RenderHeight DataFlags
 * AttachType = AttachTypeFile / AttachTypeOle
 * AttachTypeFile=%x01.00
 * AttachTypeOle=%x02.00
 * AttachPosition= INT32
 * RenderWidth=INT16
 * RenderHeight=INT16
 * DataFlags = FileDataDefault / FileDataMacBinary
 * FileDataDefault= %x00.00.00.00
 * FileDataMacBinary=%x01.00.00.00
*/

func (d *TnefDecoder) DecodeAttachmentRendData(b []byte) map[string]int {
	result := map[string]int {
		"AttachType": 0,
		"AttachPosition": 0,
		"RenderWidth":  0,
		"RenderHeight": 0,
		"DataFlags": 0,
	}

	l := len(b)

	offset := 0
	if (l-1 > offset + 2) {
		result["AttachType"] = d.leDecoder.Int(b[offset: offset+2])
	}
	offset += 2

	if (l-1 >= offset + 4) {
		result["AttachPosition"] = d.leDecoder.Int(b[offset:offset + 4])
	}
	offset+=4

	if (l-1 >= offset + 2) {
		result["RenderWidth"] = d.leDecoder.Int(b[offset:offset+2])
	}
	offset += 2

	if (l-1 >= offset + 2) {
		result["RenderHeight"] = d.leDecoder.Int(b[offset:offset+2])
	}
	offset += 2

	if (l-1 >= offset + 4) {
		result["DataFlags"] = d.leDecoder.Int(b[offset:offset+4])
	}
	offset += 4

	return result
}
