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

    itAsync("should generate an encrypted workbook", () => {
        return XlsxPopulate.fromBlankAsync()
            .then(workbook => {
                workbook.sheet(0).cell("A1").value("TEST");
                return workbook.outputAsync({ password: "SECRET" });
            })
            .then(data => {
                expect(data).toEqual(jasmine.any(Blob));
                expect(data.size).toBeGreaterThan(0);
            });
    }, 30 * 1000);

    itAsync("should generate and parse an encrypted workbook", () => {
        return XlsxPopulate.fromBlankAsync()
            .then(workbook => {
                workbook.sheet(0).cell("A1").value("TEST");
                return workbook.outputAsync({ password: "SECRET" });
            })
            .then(data => XlsxPopulate.fromDataAsync(data, { password: "SECRET" }))
            .then(workbook => {
                expect(workbook.sheet(0).cell("A1").value()).toBe("TEST");
            });
    }, 120 * 1000);
});
