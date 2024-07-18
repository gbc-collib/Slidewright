//TODO: FOR SOME REASON rIDs defined in presentation_rels are used in the slide.xml files so we need to clean the deleted an dupdate the indexdes
import JSZip, { JSZipObject } from "jszip";
const { parseStringPromise, Builder } = require('xml2js');
import fs from "fs";
import path from "path"
const outpath = './test-out'

function deleteObjectWithProperty(obj, targetKey, targetValue) {
    if (typeof obj !== 'object' || obj === null) {
        return obj;
    }

    if (Array.isArray(obj)) {
        for (let i = 0; i < obj.length; i++) {
            if (typeof obj[i] === 'object' && obj[i] !== null && obj[i][targetKey] === targetValue) {
                obj.splice(i, 1);
                i--;
            } else {
                obj[i] = deleteObjectWithProperty(obj[i], targetKey, targetValue);
            }
        }
    } else {
        for (const key in obj) {
            if (obj.hasOwnProperty(key)) {
                if (typeof obj[key] === 'object' && obj[key] !== null && obj[key][targetKey] === targetValue) {
                    delete obj[key];
                } else {
                    obj[key] = deleteObjectWithProperty(obj[key], targetKey, targetValue);
                }
            }
        }
    }

    return obj;
}

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
    public async unzipPowerPoint() {
        const powerpoint: Buffer = fs.readFileSync("./test.pptx");
        const zip = new JSZip();
        this.powerpointData = await zip.loadAsync(powerpoint);       // Extract each file from the zip
        for (const filename in zip.files) {
            const file = zip.files[filename];
            if (!file.dir) {
                // Extract the file contents
                const content: Buffer = await file.async("nodebuffer");
                const filePath = path.join(outpath, "powerpoint", filename);
                const directory = path.dirname(filePath);
                fs.mkdirSync(directory, { recursive: true });
                fs.writeFileSync(`${outpath}/powerpoint/${filename}`, content);
            }
        }
        console.log("DONE UNZIPPED");
    }


    public zipPowerPoint(): Promise<Buffer> {
        const allFilePaths: string[] = fs.readdirSync(`${outpath}/powerpoint/`, { recursive: true }) as string[];
        console.log(allFilePaths);
        const zip = new JSZip();
        for (const file of allFilePaths) {
            const stat = fs.lstatSync(`${outpath}/powerpoint/${file}`);
            if (stat.isDirectory()) {
                continue;
            }
            zip.file(file, fs.readFileSync(`${outpath}/powerpoint/${file}`));
        }
        return new Promise((resolve, reject) => {
            try {
                zip.generateNodeStream({ streamFiles: true })
                    .pipe(fs.createWriteStream(`${outpath}/powerpoint.pptx`)
                        .on("finish", () => {
                            resolve(fs.readFileSync(`${outpath}/powerpoint.pptx`));
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
    public async cleanup() {
        fs.rmSync(`${outpath}/powerpoint`, { recursive: true, force: true });
        //Possibly just delete whole ./out/
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
            relsXml.Relationships.Relationship = relsXml.Relationships.Relationship.filter(rel => {
                if (!(rel['$'].Target.includes(`slide${index}.xml`))) {
                    return true
                }
                else {
                    rId = rel['$'].Id
                    return false
                }
            }
            );
            //TODO: Parse presentation.xml and remove all references to the old rID
            //I.e. slide 1 is rID9
            //You tried but did it oncrrectly somehow
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

            for (const filename in this.powerpointData.files) {
                if (filename.includes(`ppt/slideMasters/_rels/`)) {
                    let masterSlideData = await this.powerpointData.file(filename)?.async('string');
                    let masterSlideXML = await parseStringPromise(masterSlideData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    masterSlideXML.Relationships.Relationship = masterSlideXML.Relationships.Relationship.filter((rel) => {
                        if (!(rel['$'].Id == rId)) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    })
                    const updatedSlideXML = builder.buildObject(masterSlideXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                }
                if (filename.includes(`ppt/slideMasters/slideMaster`)) {
                    let masterSlideData = await this.powerpointData.file(filename)?.async('string');
                    let masterSlideXML = await parseStringPromise(masterSlideData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    masterSlideXML = deleteObjectWithProperty(masterSlideXML, 'r:id', rId)
                    const updatedSlideXML = builder.buildObject(masterSlideXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                }
                if (filename.includes(`ppt/slides/_rels`)) {
                    let slideData = await this.powerpointData.file(filename)?.async('string');
                    let slideXML = await parseStringPromise(slideData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    slideXML.Relationships.Relationship = slideXML.Relationships.Relationship.filter((rel) => {
                        if (!(rel['$'].Id == rId)) {
                            return true;
                        }
                        else {
                            console.log("Found match, removing")
                            return false;
                        }
                    })
                    const updatedSlideXML = builder.buildObject(slideXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                }
                if (filename.includes(`ppt/slides/slide`)) {
                    let slideData = await this.powerpointData.file(filename)?.async('string');
                    let slideXML = await parseStringPromise(slideData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    //TODO: Could optimize by selecting the object we need to crawl instead of whole page
                    slideXML =
                        deleteObjectWithProperty(slideXML, 'r:embed', rId);
                    slideXML =
                        deleteObjectWithProperty(slideXML, 'r:cs', rId);

                    const updatedSlideXML = builder.buildObject(slideXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                }
                if (filename.includes(`ppt/slideLayouts/_rels`)) {
                    let slideLayoutData = await this.powerpointData.file(filename)?.async('string');
                    let layoutXML = await parseStringPromise(slideLayoutData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    layoutXML.Relationships.Relationship = layoutXML.Relationships.Relationship.filter((rel) => {
                        if (!(rel['$'].Id == rId)) {
                            return true;
                        }
                        else {
                            return false;
                        }
                    })

                    const updatedSlideXML = builder.buildObject(layoutXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                }
                if (filename.includes(`ppt/slideLayouts/slideLayout`)) {
                    let slideLayoutData = await this.powerpointData.file(filename)?.async('string');
                    let layoutXML = await parseStringPromise(slideLayoutData)
                    //console.log(masterSlideXML.Relationships.Relationship[2]);
                    layoutXML['p:sldLayout']['p:cSld'][0]['p:spTree'] =
                        deleteObjectWithProperty(layoutXML['p:sldLayout']['p:cSld'][0]['p:spTree'], 'r:id', rId);

                    const updatedSlideXML = builder.buildObject(layoutXML);
                    this.powerpointData.file(filename, updatedSlideXML)

                }
            }
            //console.log(presXML['p:presentation']['p:sldIdLst'][0]['p:sldId'])
            //Update index of all other slides to match
            relsXml.Relationships.Relationship.forEach(rel => {
                const target = rel['$'].Target;
                const match = target.match(/slide(\d+).xml/);
                if (match) {
                    const slideIndex = parseInt(match[1], 10);
                    if (slideIndex > index) {
                        const newSlideIndex = slideIndex - 1;
                        rel['$'].Target = target.replace(`slide${slideIndex}.xml`, `slide${newSlideIndex}.xml`);
                    }
                }
            });
            for (const filename in this.powerpointData.files) {
                if (filename.includes(`ppt/slides/slide`) && filename.endsWith('.xml')) {
                    const slideFile = this.powerpointData.file(filename);
                    if (slideFile) {
                        // Extract slide index from filename
                        const match = filename.match(/slide(\d+)\.xml/);
                        if (match) {
                            const slideIndex = parseInt(match[1], 10);
                            if (slideIndex > index) {
                                // Rename file in memory
                                const newFilename = `ppt/slides/slide${slideIndex - 1}.xml`;
                                this.powerpointData.file(newFilename, slideFile.async('string'));
                                this.powerpointData.remove(filename);

                                // Also rename corresponding _rels file if it exists
                                const relsFilename = `ppt/slides/_rels/slide${slideIndex}.xml.rels`;
                                const newRelsFilename = `ppt/slides/_rels/slide${slideIndex - 1}.xml.rels`;
                                const relsFile = this.powerpointData.file(relsFilename);
                                if (relsFile) {
                                    this.powerpointData.file(newRelsFilename, await relsFile.async('string'));
                                    this.powerpointData.remove(relsFilename);
                                }
                            }
                        }
                    }
                }
            }
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
    await pe.savePowerPoint()
}

test()
