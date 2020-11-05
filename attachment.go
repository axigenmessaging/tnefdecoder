/**
 * specific decoders for attachments
 */

package tnefdecoder

import (
	"strings"
//	"fmt"
)

func NewAttachment() *Attachment {
	return &Attachment{
		TnefObject: &TnefObject {},
		Filename: "",
		Data: []byte{},
		DecodedAttRendData: make(map[string]int),
	}
}


type Attachment struct {
	*TnefObject
	Filename string
	Data []byte

	// custom attributes that needs specific decoders
	DecodedAttRendData map[string]int
}


/**
 * set attachment filename
 */
func (a *Attachment) SetFilename(t string) {
	a.Filename = t
}

/**
 * set attachment data
 */
func (a *Attachment) SetData(d []byte) {
	a.Data = d
}

/**
 * get attachment body
 */
func (a *Attachment) GetData() []byte {
	if len(a.Data) == 0 {
		// if the data is not already set (custom, using SetData, or previously requested) try to extract the data from attributes
		attr  := a.GetAttribute(AttAttachData, "mapped")

		if attr != nil {
			a.SetData(attr.Data)
		}
	}

	return a.Data
}


func (a *Attachment) GetFilename() string {
	var attr *Attribute

	/*
	attr  = a.GetAttribute(AttAttachTransportFilename, "mapped")

	if attr != nil {
		fmt.Println("Transport Name:", attr.GetStringValue())
	}
	*/

	if a.Filename == "" {
		// if the Filename is not already set (custom, using SetFilename, or previously requested) try to extract the data from attributes
		attr  = a.GetAttribute(AttAttachTitle, "mapped")
		if attr != nil && attr.GetStringValue() != "" {
			a.SetFilename(attr.GetStringValue())
		}
	}

	if a.Filename == "" {
		// if the Filename is not already set (custom, using SetFilename, or previously requested) try to extract the data from attributes
		attr  = a.GetAttribute(MapiPidTagDisplayName, "mapi")
		if attr != nil && attr.GetStringValue() != "" {
			a.SetFilename(attr.GetStringValue())
		}
	}

	if a.Filename == "" {
		// if the Filename is not already set (custom, using SetFilename, or previously requested) try to extract the data from attributes
		attr  = a.GetAttribute(MapiPidTagAttachFilename, "mapi")
		if attr != nil && attr.GetStringValue() != "" {
			a.SetFilename(attr.GetStringValue())
		}
	}

	if a.Filename == "" {
		extension := ""
		// if the Filename is not already set (custom, using SetFilename, or previously requested) try to extract the data from attributes
		attr  = a.GetAttribute(MapiPidTagAttachExtension, "mapi")
		if attr != nil && attr.GetStringValue() != "" {
			extension = attr.GetStringValue()
		}
		a.SetFilename("unkown"+extension)
	}


	if strings.Count(a.Filename, "?") != 0 || len(a.Filename) > 255 {
		// the filename has invalid characters or is too long
		attr  = a.GetAttribute(MapiPidTagAttachFilename, "mapi")
		if attr != nil && attr.GetStringValue() != "" {
			a.SetFilename(attr.GetStringValue())
		}
	}

	return a.Filename
}

/**
 * check if the attachment has a reference in html as cid
 * @param  {[type]} a *Attachment)  IsMimeRelated( [description]
 * @return {[type]}   [description]
 */
 func (a *Attachment) HasCID() (bool) {
	cid := a.GetCID()
	// is not embeded -> the attachment data is kept into attAttachData attribute
	return cid != ""
}

func (a *Attachment) GetCID() string {
	attContentIdAttr := a.GetAttribute(MapiPidTagAttachContentId, "mapi");
	if attContentIdAttr == nil {
		return ""
	}
	return attContentIdAttr.GetStringValue()
}

/*
 *
 * AttachTypeFile= 1
 * AttachTypeOle = 2
 * if we cannot find the information return 0
 */
func (a *Attachment) GetRenderType() int {
	if r, ok := a.DecodedAttRendData["AttachType"]; ok {
		return r
	}
	return 0
}


