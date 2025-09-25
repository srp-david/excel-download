package com.lannstark.excel;

import java.io.IOException;
import java.io.OutputStream;
import java.util.List;

public interface ExcelFile<T> {
	void write(OutputStream stream) throws IOException;
    void addRows(List<T> data);
    default List<Object> read() { return List.of(); }
    default <T> List<T> read(Class<T> type) { return null;};
}
