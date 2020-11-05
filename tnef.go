package tnefdecoder

import (
	"strings"
	rtf "rtfconverter"

)

/**
 * may be any value type; when you use it, usualy you know what type has the attribute value so you can do a type assertion
 */
 type GenericValue interface{}

 type TnefObject struct {
	// tnef attributes (extracted from TNEFVersion, OEMCodePage, MessageAttribute) + MAPI attributes (extracted from MessageProps)
	Attributes []*Attribute

	// attachments extracted from AttachData
	Attachments []*Attachment

	TextBody []byte
	HtmlBody []byte
}


/**
 * get an attribute by MAPI ID/TNEF Attribute ID
 */
func (t *TnefObject) GetAttribute(attrId int, attrType string) (attr *Attribute) {
	for _, attr = range t.Attributes {
		if attr.Id == attrId && attr.Type == attrType {
			return
		}
	}
	return nil
}

func (t *TnefObject) GetHtmlBody() []byte {
	if t.HtmlBody == nil {
	   attr := t.GetAttribute(MapiPidTagBodyHtml, "mapi")

	   if attr != nil {
		   t.HtmlBody = []byte(attr.GetStringValue())
	   } else {
		   t.HtmlBody = []byte("")
	   }
   }

   return t.HtmlBody
}

func (t *TnefObject) SetHtmlBody(v []byte)  {
	t.HtmlBody = v
}


func (t *TnefObject) GetTextBody() []byte {
   if t.TextBody == nil {
	   attr := t.GetAttribute(MapiPidTagBody, "mapi")

	   if attr != nil {
		   t.TextBody = []byte(attr.GetStringValue())
	   } else {
		   t.TextBody = []byte("")
	   }
   }

   return t.TextBody
}

func (t *TnefObject) SetTextBody(v []byte)  {
   t.TextBody = v
}

/**
* return message class
* If the value of the attMessageClass or attOriginalMessageClass attribute begins with the string "Microsoft Mail v3.0 ",
* the TNEF Reader MUST ignore the "Microsoft Mail v3.0 " prefix when attempting to match the value of the attMessageClass or attOriginalMessageClass
*/

func (t *TnefObject) GetMessageClass() string {
   attr := t.GetAttribute(AttMessageClass, "mapped")
   messageClass := ""
   if attr != nil {
	   messageClass = attr.GetStringValue()
   }
   if messageClass == "" {
	   attr = t.GetAttribute(AttOriginalMessageClass, "mapped")
	   if attr != nil {
		   messageClass = attr.GetStringValue()
	   }
   }
   messageClass = strings.TrimPrefix(messageClass, "Microsoft Mail v3.0 ")

   return messageClass
}


/**
 * Value Meaning
 * 0x00000000 Undefined body
 * 0x00000001 Plain text body
 * 0x00000002 Rich Text Format (RTF) compressed body
 * 0x00000003 HTML body
 * 0x00000004 Clear-signed body
 */
 func (t *TnefObject) GetBodyFormat() int {
	attr := t.GetAttribute(MapiPidTagNativeBody, "mapi")
	if attr != nil{
		return attr.GetIntValue()
	}
	return 0
}


 /**
  *  decode compressed RTF from MapiPidTagRtfCompressed
  *  if exists, will rewrite TNEF object TEXT / HTML value or add an attachment
  */
func (t *TnefObject) DecodeRtf() {
	rtfContentAttr := t.GetAttribute(MapiPidTagRtfCompressed, "mapi")
	if rtfContentAttr != nil && len(rtfContentAttr.GetBinaryValue()) > 0 {

		/**
		*  Try to decompress RTF
		*  @TODO: to check if we can use MapiPidTagNativeBody or MapiPidTagRtfInSync
		*/
		data, err := rtf.Decompress(rtfContentAttr.GetBinaryValue())


		if err == nil {

			attachRtf := true

			c := rtf.NewConverter()
			c.SetBytes(data)
			html, err := c.Convert("html")
			if err == nil && html !=nil && len(html) > 0 {
				attachRtf = false
				t.SetHtmlBody(html)
			} else {
				text, err := c.Convert("text")
				if err == nil && text !=nil && len(text) > 0 {
					attachRtf = false
					t.SetTextBody(text)
				}
			}

			if attachRtf {
				// add the file as attachment
				attachment := NewAttachment()
				attachment.SetFilename("message.rtf")
				attachment.SetData(data)
				t.Attachments = append(t.Attachments, attachment)
			}
		}
	}
}

