

type MasterInfo = {
    masterName: string,
    level1Chapter: {
        list: {
            zh: string,
            en: string
        }[],
        activeIndex: number
    },
    level2Chapter?: {
        list: {
            zh: string,
        }[],
        activeIndex: number
    }
    level3Chapter?: {
        list: {
            zh: string,
        }[],
        activeIndex: number
    }
}