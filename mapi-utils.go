/**
 * mapi attributes value utils (conversions, etc)
 */
package tnefdecoder



/**
; Scalars – Types Int16, Int32, Flt32, Flt64, Currency, AppTime,
; Bool, Int64, Systime, CLSID
; The data for the particular property is written to the stream and if necessary,
; padded with bytes (which SHOULD be zero) to achieve a multiple of 4-bytes in length.
PropertyScalarContent = MsgPropertyContent [PropertyPad]

; Multi-value Scalars – Types MVInt16, MVInt32, MVFlt32, MVFlt64, MVCurrency,
; MVAppTime, MVInt64, MVSystime, MVCLSID
; The number of values for the property is written to the stream as a 4-byte
; value, then the data for each value is written to the stream and if need
; be, padded with bytes (which SHOULD be zero) to achieve a multiple of 4 bytes in length.
PropertyMultiScalarContent = PropertyContentCount *PropertyScalarContent
PropertyContentCount = UINT32

; Variable-length – Types Unicode, String8, Object, Binary.
; These are handled as a special case of Multi-Variable-Length with the number of values=1.
; Multi-Variable-length – Types MVUnicode, MVString8, MVBinary
; The number of values for the property is written to the stream as a 4-byte value. ; Then,
for each value, the size of the property is written to the stream as a
; 4-byte value, then the data for the property is written to the stream and, if
; necessary, padded with zero bytes to achieve a multiple of 4 bytes in length.
PropertyMultiVariableContent = MsgPropertyCount 1*PropertyVariableContent
MsgPropertyCount = UINT32

PropertyVariableContent = MsgPropertySize MsgPropertyContent [PropertyPad]
MsgPropertySize = UINT32


bytesLenth - the length of the value  (in bytes) - if bytesLength < 0 => variable length
isMultiValue - is an array of values of the same fix bytes length

 */
func GetTypeSize(mapiType int) (bytesLength int, isMultiValue bool) {
	bytesLength = 0

	/**
	*  multiple values are Variable size; a COUNT field followed by that many mapiType (non variable mapi type) values.
	*/
	isMultiValue = (mapiType & 0x1000) > 0

	switch (mapiType) {
		case MapiTypeNull:
			bytesLength = 0
		case MapiTypeInt16, MapiTypeMVInt16, MapiTypeBoolean:
			bytesLength = 2
		case MapiTypeInt32, MapiTypeMVInt32, MapiTypeFlt32, MapiTypeMVFlt32:
			bytesLength = 4
		case MapiTypeFlt64, MapiTypeMVFlt64, MapiTypeCurrency, MapiTypeMVCurrency, MapiTypeAppTime, MapiTypeMVAppTime, MapiTypeInt64, MapiTypeMVInt64, MapiTypeSystime, MapiTypeMVSystime:
			bytesLength = 8
		case MapiTypeCLSID, MapiTypeMVCLSID:
			bytesLength = 16
		case MapiTypeString8, MapiTypeMVString8, MapiTypeUnicode, MapiTypeMVUnicode, MapiTypeBinary, MapiTypeMVBinary, MapiTypeObject:
			// is variable size, should be read from the value
			bytesLength = -1
			isMultiValue = true // force to multivalues because the Variable-length type is treat it as a special case for PropertyMultiVariableContent with MsgPropertyCount=1
	}

	return
}

/**
 * decode mapi PropertyMultiScalarContent property  value
 * Multi-value Scalars – Types MVInt16, MVInt32, MVFlt32, MVFlt64, MVCurrency,
 * MVAppTime, MVInt64, MVSystime, MVCLSID
 * The number of values for the property is written to the stream as a 4-byte
 * value, then the data for each value is written to the stream and if need
 *  be, padded with bytes (which SHOULD be zero) to achieve a multiple of 4 bytes in length.
 * PropertyMultiScalarContent = PropertyContentCount *PropertyScalarContent
 * PropertyContentCount = UINT32
 *
 */

func GetPropertyMultiScalarValues(b []byte, dataType int) [][]byte {
	result := [][]byte{}
	offset := 0
	valueLength := 0
	padd := 0
	leReader := new(LittleEndianDecoder)
	value := []byte{}

	c := int(leReader.Uint32(b[offset : offset + 4]))
	offset += 4
	for i:=0; i< c; i++ {
		valueLength, _ = GetTypeSize(dataType) // do not count the padding
		value = b[offset:offset+valueLength]
		offset += valueLength

		if padd = 4 - valueLength % 4; padd < 4 {
			offset += padd
		}

		result = append(result, value)
	}

	return result
}


/**
 * decode MAPI Variable-length & Multi-Variable-length  and
 *
 * Variable-length – Types Unicode, String8, Object, Binary.
 * These are handled as a special case of Multi-Variable-Length with the number of values=1.
 * Multi-Variable-length – Types MVUnicode, MVString8, MVBinary
 * The number of values for the property is written to the stream as a 4-byte value. ; Then,
 * for each value, the size of the property is written to the stream as a
 * 4-byte value, then the data for the property is written to the stream and, if
 * necessary, padded with zero bytes to achieve a multiple of 4 bytes in length.
 * PropertyMultiVariableContent = MsgPropertyCount 1*PropertyVariableContent
 * MsgPropertyCount = UINT32
 * The size of the property is written to the stream as a 4-byte value, then the data
 * for the property is written to the stream and if necessary, padded with zero bytes
 * to achieve a multiple of 4 bytes in length. The size includes the Interface
 * Identifier at the beginning of the value stream for an object but does not include
 * the padding bytes.
 */

func GetPropertyMultiVariableValues(b []byte) [][]byte {
	result := [][]byte{}
	offset := 0
	valueLength := 0
	padd := 0
	leReader := new(LittleEndianDecoder)
	value := []byte{}

	c := int(leReader.Uint32(b[offset : offset + 4]))
	offset += 4
	for i:=0; i< c; i++ {
		valueLength = int(leReader.Uint32(b[offset : offset + 4])) // do not count the padding
		offset += 4
		value = b[offset:offset+valueLength]
		offset += valueLength

		if padd = 4 - valueLength % 4; padd < 4 {
			offset += padd
		}

		result = append(result, value)
	}

	return result
}

/**
 * decode MAPI string8 scalar value
 */
func MapiDecodeString8(b []byte) string {
	var result string
	items := GetPropertyMultiVariableValues(b)

	if (len(items) > 0) {
		result = string(items[0])
	}

	return result
}

/**
 * decode MAPI string8 Multi-value Scalars value
 */
func MapiDecodeString8Array(b []byte) []string {
	var result []string
	items := GetPropertyMultiVariableValues(b)

	for i:=0; i < len(items); i++ {
		result = append(result, string(items[i]))
	}

	return result
}

// MapiTypeUnicode
func MapiDecodeUnicode(b []byte) string {
	var result string
	leReader := new(LittleEndianDecoder)
	items := GetPropertyMultiVariableValues(b)
	if (len(items) > 0) {
		result = leReader.Utf16(items[0])
	}

	return result
}

//MapiTypeMVUnicode
func MapiDecodeUnicodeArray(b []byte) []string {
	var result []string
	leReader := new(LittleEndianDecoder)
	items := GetPropertyMultiVariableValues(b)

	for i:=0; i < len(items); i++ {
		result = append(result, leReader.Utf16(items[i]))
	}

	return result
}

/**
 * convert MAPI all int types scalars as int
 * if need it, extend with new function to decode specfic int value (INT16, INT32, eetc)
 */

func MapiDecodeInt(b []byte, dataType int) int {
	v := 0

	leReader := new(LittleEndianDecoder)

	switch (dataType) {
		case MapiTypeInt16:
			v = int(leReader.Int16(b))
		case MapiTypeInt32:
			v = int(leReader.Int32(b))
		case MapiTypeInt64, MapiTypeSystime, MapiTypeCurrency:
			v = int(leReader.Int64(b))
		default:
			v = int(leReader.Int(b))
	}

	return v
}

/**
	dataType may be  MapiTypeMVInt16, 	MapiTypeMVInt32 , 	MapiTypeMVInt64, MapiTypeMVSystime, MapiTypeMVCurrency
*/
func MapiDecodeIntArray(b []byte, dataType int) []int {
	items  := GetPropertyMultiScalarValues(b, dataType)
	result := make([]int, len(items))
	for i, item := range items {
		result[i] = MapiDecodeInt(item, dataType)
	}

	return result
}

// MapiTypeFlt32
func MapiDecodeFloat32(b []byte) float32 {
	leReader := new(LittleEndianDecoder)
	return leReader.Float32(b)
}

// MapiMVTypeFlt32
func MapiDecodeFloat32Array(b []byte) []float32 {
	items  := GetPropertyMultiScalarValues(b, MapiTypeMVFlt64)
	result := make([]float32,len(items))
	for i, item := range items {
		result[i] = MapiDecodeFloat32(item)
	}
	return result
}

// MapiTypeAppTime, MapiTypeFlt64
func MapiDecodeFloat64(b []byte) float64 {
	leReader := new(LittleEndianDecoder)
	return leReader.Float64(b)
}

// MapiTypeMVAppTime, MapiMVTypeFlt64
func MapiDecodeFloat64Array(b []byte) []float64 {
	items  := GetPropertyMultiScalarValues(b, MapiTypeMVFlt64)
	result := make([]float64, len(items))
	for i, item := range items {
		result[i] = MapiDecodeFloat64(item)
	}
	return result
}

func MapiDecodeBoolean(b []byte) bool {
	leReader := new(LittleEndianDecoder)
	return leReader.Boolean(b)
}

func MapiDecodeObject(b []byte) []byte {
	var result []byte
	offset := 0
	leReader := new(LittleEndianDecoder)
	noOfValues := int(leReader.Uint32(b[offset:offset + 4])) // should be always 1
	offset += 4
	for i := 0; i < noOfValues; i++ {
		bytesLength := int(leReader.Uint32(b[offset:offset + 4]))
		offset += 4

		result = b[offset:offset+bytesLength]
		offset += bytesLength

		if padd := 4 - bytesLength % 4; padd < 4  {
			offset += padd
		}
	}
	return result
}

func MapiDecodeBinary(b []byte) []byte {
	var result []byte

	items := GetPropertyMultiVariableValues(b)
	if (len(items) > 0) {
		result = items[0]
	}
	return result
}

func MapiDecodeBinaryArray(b []byte) [][]byte {
	return GetPropertyMultiVariableValues(b)
}
