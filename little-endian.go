/**
 * little endian decoder/reader
 */

package tnefdecoder

import (
	"bytes"
	"encoding/binary"
	"unicode/utf16"
)

/**
 * little endian reader
 */
type LittleEndianDecoder struct {

}

// read an int
func (r *LittleEndianDecoder) Int(data []byte) int {
	var num int
	var n uint
	for _, b := range data {
		num += (int(b) << n)
		n += 8
	}
	return num
}


func (c *LittleEndianDecoder) String(b []byte) (string) {
	var v string
	buf := bytes.NewReader(b)
	binary.Read(buf, binary.LittleEndian, &v)
	return v
}


// Int = Int32
func (c *LittleEndianDecoder) Int32(b []byte) (int32) {
	var v int32
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}
// UInt = UInt32
func (c *LittleEndianDecoder) Uint32(b []byte) (uint32) {
 	return binary.LittleEndian.Uint32(b)
}

// int64
func (c *LittleEndianDecoder) Int64(b []byte) (int64) {
	return int64(c.Int(b))
}

// uint64
func (c *LittleEndianDecoder) Uint64(b []byte) (uint64) {
	return binary.LittleEndian.Uint64(b)
}

func (c *LittleEndianDecoder) Int16(b []byte) (int16) {
	return int16(c.Int(b))
}

func (c *LittleEndianDecoder) Uint16(b []byte) (uint16) {
	return binary.LittleEndian.Uint16(b)
}

func (c *LittleEndianDecoder) Float32(b []byte) (float32) {
	var v float32
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianDecoder) Float64(b []byte) (float64) {
	var v float64
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

func (c *LittleEndianDecoder) Boolean(b []byte) (bool) {
	var v bool
	buf := bytes.NewReader(b)
 	binary.Read(buf, binary.LittleEndian, &v)
	return v
}

// read utf16 little endian
func (c *LittleEndianDecoder) Utf16(content []byte) (convertedStringToUnicode string) {
	tmp := []uint16{}

	bytesRead := 0
	for {
		tmp = append(tmp, binary.LittleEndian.Uint16(content[bytesRead:]))
		bytesRead += 2

		convertedStringToUnicode = string(utf16.Decode(tmp));

		if (len(content) <= bytesRead) {
			break
		}
	}
	return
}
