import pprint


def main():
    species = ["cats", "dogs", "pugs"]
    matrixes = ["livers", "kidney", "baalls"]
    elements = ["copper", "mustar"]
    range_dict = {
        "range_lo": "< 0.1",
        "range_hi": "> 10"
    }

    pp = pprint.PrettyPrinter(indent=4)

    a_dict = {}

    for s in species:
        if s not in a_dict:
            a_dict[s] = {}
        s_dict = a_dict[s]

        for m in matrixes:
            if m not in s_dict:
                s_dict[m] = {}
            m_dict = s_dict[m]

            for e in elements:
                if e not in m_dict:
                    m_dict[e] = {}
                m_dict[e] = range_dict

    pp.pprint(a_dict)


if __name__ == "__main__":
    main()
