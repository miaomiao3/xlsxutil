package xlsxutil

import "reflect"

// SliceSetter add new element to slice using reflection
// suitable for
//		slice of struct
//		slice of struct_ptr
type SliceSetter struct {
	isElementPtr bool
	elementType  reflect.Type
	sliceValue   reflect.Value // value of input data
	newSlice     reflect.Value // buffer to new element
	optionMap    map[string]*xlsOption
}

func NewSliceSetter(data interface{}) (*SliceSetter, error) {
	// validate
	elementType, err := validateDataInput(data)
	if err != nil {
		return nil, err
	}
	setter := &SliceSetter{}
	setter.sliceValue = reflect.ValueOf(data).Elem()
	setter.newSlice = reflect.MakeSlice(reflect.SliceOf(elementType), 0, 10)

	if elementType.Kind() == reflect.Ptr {
		setter.isElementPtr = true
		elementType = elementType.Elem()
	}
	setter.elementType = elementType

	elementValueSample := reflect.New(elementType).Elem()
	_, optionMap := getStructOptions(elementValueSample)
	setter.optionMap = optionMap

	return setter, nil
}

func (setter *SliceSetter) AddElement(valueMap map[string]string) {
	elem := newElement(setter.elementType, valueMap, setter.optionMap)
	if setter.isElementPtr {
		setter.newSlice = reflect.Append(setter.newSlice, (*elem).Addr())
	} else {
		setter.newSlice = reflect.Append(setter.newSlice, *elem)
	}
}

// Update reflect set the data with slice elements
func (setter *SliceSetter) Update() {
	setter.sliceValue.Set(setter.newSlice)
}
