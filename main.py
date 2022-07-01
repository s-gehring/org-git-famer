import json
import os
import pathlib

import git
import gitfame
import xlsxwriter
from dotenv import load_dotenv
from git import Repo
from github import Github, Organization, Repository
from xlsxwriter.worksheet import Worksheet

load_dotenv()

GITHUB_TOKEN = os.getenv("GITHUB_TOKEN")
ORGANIZATION = os.getenv("GITHUB_ORGANIZATION")

github: Github = Github(GITHUB_TOKEN)

toIgnoreExtensions = [
    "eot", "ico", "jpg", "png", "ttf", "woff", "woff2", "svg", "gzip", "zip", "crt", "pub", "_None_ext", "key", "otf"
]
toIgnore = list(map(lambda x: "(.+\\.{0})".format(x), toIgnoreExtensions)) + ["(package-lock\\.json)", "(yarn\\.lock)",
                                                                              "(.*/node_modules/.*)"]
toIgnorePattern = "(" + "|".join(toIgnore) + ")"
print(toIgnorePattern)

broestech: Organization = github.get_organization(ORGANIZATION)

print("Getting repos for organization {0}.".format(broestech.company))

print(broestech.repos_url)

repos = broestech.get_repos(type="all", sort="updated", direction="desc")

temp_dir = pathlib.Path("repos")
if not temp_dir.exists():
    temp_dir.mkdir()

print("Found {0} repositories in total.".format(repos.totalCount))
homeDir = pathlib.Path.cwd()


def getResultForRepo(repository: Repository, ignore_cache=False):
    print("Currently processing {0}.".format(repository.name))
    clone_url = "git@github.com:" + broestech.login + "/" + repository.name + ".git"
    repo_folder = temp_dir / repository.name
    cache_file = repo_folder / "gitfame-results.json"
    if repo_folder.exists():
        existing_repo = git.Repo(repo_folder)
        existing_repo.remote("origin").pull()
    else:
        repo_folder.mkdir()
        Repo.clone_from(clone_url, repo_folder)
    if cache_file.exists():
        if ignore_cache:
            print("Ignoring cached file and removing it.")
            cache_file.unlink()
        else:
            with open(str(cache_file)) as json_file:
                print("Restoring cached results for {0}.".format(repository.name))
                return json.load(json_file)

    os.chdir(str(repo_folder))
    result = json.loads(gitfame.main(
        ["--format=json", "-e", "-w", "-M", "-C", "--excl=" + toIgnorePattern]))
    os.chdir(str(homeDir))

    with open(str(cache_file.absolute()), "w") as json_file:
        json.dump(result, json_file)
        print("Writing results to {0}.".format(str(cache_file.absolute())))
    return result


def writeExcelHead(worksheet, name: str):
    worksheet.write("A1", name)
    worksheet.write("A2", "Who? (Email)")
    worksheet.write("B2", "Lines of Code")
    worksheet.write("C2", "Commits")
    worksheet.write("D2", "Files")


def write(colChar: str, rowNumber: int, content: str, worksheet: Worksheet):
    worksheet.write(colChar + str(rowNumber), content)


def get_worksheet_name(fullname: str) -> str:
    return fullname[:30]


def serializeToCsv(to_serialize):
    target = xlsxwriter.Workbook("results.xlsx")
    for one_repo_result in to_serialize:
        worksheet = target.add_worksheet(get_worksheet_name(one_repo_result["name"]))
        writeExcelHead(worksheet, one_repo_result["name"])
        i = 3
        for datapoint in one_repo_result["results"]["data"]:
            write("A", i, datapoint[0], worksheet)  # Name
            write("B", i, datapoint[1], worksheet)  # LoC
            write("C", i, datapoint[2], worksheet)  # Commits
            write("D", i, datapoint[3], worksheet)  # Files
            i += 1
    worksheet = target.add_worksheet("Summary")
    worksheet.write("A1", "Summary")
    worksheet.write("A2", "Project")
    worksheet.write("B2", "LoC")
    worksheet.write("C2", "Commits")
    worksheet.write("D2", "Files")
    i = 3
    for one_repo_result in to_serialize:
        write("A", i, one_repo_result["name"], worksheet)
        write("B", i, one_repo_result["results"]["total"]["loc"], worksheet)
        write("C", i, one_repo_result["results"]["total"]["commits"], worksheet)
        write("D", i, one_repo_result["results"]["total"]["files"], worksheet)
        i += 1
    target.close()


results: [dict] = []
for repo in repos:
    results.append({
        "name": repo.name, "results": getResultForRepo(repo)
    })
serializeToCsv(results)
