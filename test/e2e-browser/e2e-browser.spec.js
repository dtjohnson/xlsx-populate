"use strict";

describe("e2e-browser", () => {
    it("should expose XlsxPopulate globally", () => {
        expect(XlsxPopulate).toBeDefined();
    });

    itAsync("should generate a workbook", () => {
        return XlsxPopulate.fromBlankAsync()
            .then(workbook => {
                workbook.sheet(0).cell("A1").value("TEST").style("fontColor", "red");
                return workbook.outputAsync();
            })
            .then(data => {
                expect(data).toEqual(jasmine.any(Blob));
                expect(data.size).toBeGreaterThan(0);
            });
    });
});
