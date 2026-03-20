type TextItem = {
    text: string,
    // 是否需要项目符号(标题下的内容不需要)
    hasBullet: boolean,
    // 子内容,这里的内容，就需要使用项目符号了
    subContent: TextItem[]
}

type ImageContent = {
    "id": string
    "uri": string
}

type Content = {
    // 标题，取上级章节的中文名
    title: string,
    // 标题英文，取上级章节的英文名
    titleEn?: string,
    // 所属章节序号，从0开始，方便定位
    index: number[]
    // 内容类型, 相邻text类型需要聚合
    type: 'text' | 'image',
    // 普通文本，可能包含子文本，子文本需要使用项目符号，内容过长需要分页
    text?: TextItem[],
    // 图片，图片需要独立一页，高度填满，高度自适应
    image?: ImageContent
}

type Chapter = {
    // 编号，例如： 1, 1.1, 1.1.1
    code: string,
    // 级别，例如 1, 2, 3
    level: number,
    // 章节名称
    chapterName: string,
    // 章节英文名称(级别为1的章节，肯定有，其他级别不一定有)
    chapterEnName?: string,
    // 章节内容（如果内容长度超过9个字符，则判断为内容）
    content?: Content[],
    // 子章节（不是章节内容，那么就是子章节）
    subChapter?: Chapter[]
}

type Course = {
    // 课程名称
    courseName: string,
    // 课程英文名称
    courseEnName: string
    // 章节
    chapters: Chapter[]
}