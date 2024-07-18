//TODO: FOR SOME REASON rIDs defined in presentation_rels are used in the slide.xml files so we need to clean the deleted an dupdate the indexdes
import JSZip, { JSZipObject } from "jszip";
const { parseStringPromise, Builder } = require('xml2js');
import fs from "fs";
const outpath = './test-out'

export class PowerPointEditor {
    public powerpointData: JSZip = new JSZip();
    public async loadPowerPoint() {
        const powerpoint: Buffer = fs.readFileSync("./test.pptx");
        const zip = new JSZip();
        this.powerpointData = await zip.loadAsync(powerpoint);       // Extract each file from the zip
    }
    async savePowerPoint() {

        return new Promise((resolve, reject) => {
            try {
                if (!this.powerpointData) {
                    throw new Error("No powerpoint data")
                }
                this.powerpointData.generateNodeStream({ streamFiles: true })
                    .pipe(fs.createWriteStream(`${outpath}/test-powerpoint.pptx`)
                        .on("finish", () => {
                            resolve('Done');
                        }));
            }
            catch (error) {
                reject(error);
            }
        });

    }

    public writePowerPoint(powerpoint: Buffer, filePath: string) {
        fs.writeFileSync(filePath, powerpoint)

    }

    public async deleteSlide(index: number) {
        const slideXmlPath = `ppt/slides/slide${index}.xml`;
        this.powerpointData.remove(slideXmlPath)
        this.powerpointData.remove(`ppt/slides/_rels/slide${index}.xml.rels`)
        const presPath = 'ppt/presentation.xml';
        const relsPath = 'ppt/_rels/presentation.xml.rels';
        const relsData = await this.powerpointData.file(relsPath)?.async('string');
        if (relsData) {
            var relsXml = await parseStringPromise(relsData);
            //Delete all refs to slide we removed
            var rId = 0
            relsXml.Relationships.Relationship.forEach(rel => {
                if (!(rel['$'].Target.includes(`slide${index}.xml`))) {
                    return true
                }
                else {
                    rId = rel['$'].Id
                    return false
                }
            }
            );
            var presentationData = await this.powerpointData.file(presPath)?.async('string');
            var presXML = await parseStringPromise(presentationData)
            presXML['p:presentation']['p:sldIdLst'][0]['p:sldId'] = presXML['p:presentation']['p:sldIdLst'][0]['p:sldId'].filter((rel) => {
                if (!(rel['$']['r:id'] == rId)) {
                    return true;
                }
                else {
                    console.log("ID was found removing")
                    return false
                }
            })
            var builder = new Builder();
            const updatedPresXML = builder.buildObject(presXML);
            this.powerpointData.file(presPath, updatedPresXML);
            // Convert back to XML string and update in zip
            var builder = new Builder();
            const updatedRelsXml = builder.buildObject(relsXml);
            //const updatedRelsXml = xml2js.Builder(relsXml, { compact: true });
            this.powerpointData.file(relsPath, updatedRelsXml);

        } else {
            throw new Error('Failed to read presentation.xml.rels');
        }
    }
}


const test = async function() {
    const pe = new PowerPointEditor();
    await pe.loadPowerPoint()
    await pe.deleteSlide(1)
    await pe.deleteSlide(1)
    await pe.savePowerPoint()
}

test()
