package tnefdecoder

import (
	"strings"
)

/**
  * MAPI attributes & TNEF Attributes
  */
  type Attribute struct {
	Type string // mapi (decoded mapi property) or mapped (message or attachment mapped property)

	// level: att
	Level int
	Id int
	DataType int
	Data []byte  // generated only for mapped properties

	Checksum []int //UINT16
	GUID string
	PropMapValueType int // 0 = int, 1 = string
	PropMapValue GenericValue
}

func (a *Attribute) GetStringValue() string {
	v := ""
   switch a.DataType {
	   case MapiTypeString8:
		   v = MapiDecodeString8(a.Data)
	   case MapiTypeUnicode:
		   v = MapiDecodeUnicode(a.Data)
	   default:
		   v = string(a.Data)
   }
   return strings.TrimSuffix(v, "\x00")
}

func (a *Attribute) GetStringValueArray() []string {
   result := []string{""}
  switch a.DataType {
	  case MapiTypeMVString8:
		   result = MapiDecodeString8Array(a.Data)
	  case MapiTypeMVUnicode:
		   result = MapiDecodeUnicodeArray(a.Data)
  }
  return result
}

func (a *Attribute) GetIntValue() int {
   return MapiDecodeInt(a.Data, a.DataType)
}

func (a *Attribute) GetBoolValue() bool {
   return MapiDecodeBoolean(a.Data)
}

func (a *Attribute) GetObjectValue() []byte {
   return MapiDecodeObject(a.Data)
}

func (a *Attribute) GetBinaryValue() []byte {
   return MapiDecodeBinary(a.Data)
}

func (a *Attribute) GetBinaryValueArray() [][]byte {
   return MapiDecodeBinaryArray(a.Data)
}
